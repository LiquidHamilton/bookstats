import { useEffect, useRef, useState } from "react";
import { Camera, CheckCircle2, Keyboard, ScanLine, X } from "lucide-react";
import { normalizeIsbn } from "@bookstats/domain";

type BarcodeResultLike = { rawValue?: string };
type BarcodeDetectorLike = { detect(source: CanvasImageSource): Promise<BarcodeResultLike[]> };
type BarcodeDetectorCtor = new (options?: { formats?: string[] }) => BarcodeDetectorLike;

interface Props {
  onDetected: (isbn: string) => void;
  onClose: () => void;
}

export function BarcodeScanner({ onDetected, onClose }: Props) {
  const videoRef = useRef<HTMLVideoElement>(null);
  const streamRef = useRef<MediaStream | undefined>(undefined);
  const controlsRef = useRef<{ stop(): void } | undefined>(undefined);
  const runningRef = useRef(true);
  const [manual, setManual] = useState("");
  const [status, setStatus] = useState("Starting camera…");
  const [cameraReady, setCameraReady] = useState(false);
  const [detectorSupported, setDetectorSupported] = useState(true);

  function handleBarcode(raw: string): boolean {
    const value = normalizeIsbn(raw);
    if (isBookIsbn(value)) {
      runningRef.current = false;
      controlsRef.current?.stop();
      setStatus(`ISBN ${value} found.`);
      onDetected(value);
      return true;
    }
    if (value) setStatus(`Barcode ${value} is not an ISBN-10/ISBN-13 book code. Keep scanning.`);
    return false;
  }

  useEffect(() => {
    let timer: number | undefined;
    async function start() {
      if (!navigator.mediaDevices?.getUserMedia) { setDetectorSupported(false); setStatus("Camera access is not available here. Enter the ISBN below instead."); return; }
      const ctor = (globalThis as typeof globalThis & { BarcodeDetector?: BarcodeDetectorCtor }).BarcodeDetector;
      try {
        if (ctor) {
          const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: "environment" }, width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false });
          streamRef.current = stream;
          const video = videoRef.current;
          if (!video) return;
          video.srcObject = stream;
          await video.play();
          setCameraReady(true); setStatus("Point the camera at the ISBN barcode on the back of the book.");
          const detector = new ctor({ formats: ["ean_13", "ean_8", "upc_a", "upc_e", "code_128"] });
          const scan = async () => {
            if (!runningRef.current || !videoRef.current) return;
            try {
              const results = await detector.detect(videoRef.current);
              for (const result of results) if (handleBarcode(result.rawValue ?? "")) return;
            } catch { /* transient video frames can fail while the camera warms up */ }
            if (runningRef.current) timer = window.setTimeout(() => void scan(), 220);
          };
          void scan();
          return;
        }

        // Safari and other browsers without BarcodeDetector use ZXing. It is loaded
        // only when the scanner opens so ordinary BookStats sessions do not pay the
        // barcode-decoder bundle cost.
        setDetectorSupported(true);
        setStatus("Starting camera scanner…");
        const { BrowserMultiFormatReader } = await import("@zxing/browser");
        const reader = new BrowserMultiFormatReader(undefined, { delayBetweenScanAttempts: 180 });
        const video = videoRef.current;
        if (!video) return;
        const controls = await reader.decodeFromConstraints(
          { video: { facingMode: { ideal: "environment" }, width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false },
          video,
          (result) => {
            if (!result || !runningRef.current) return;
            handleBarcode(result.getText());
          }
        );
        controlsRef.current = controls;
        setCameraReady(true);
        setStatus("Point the camera at the ISBN barcode on the back of the book.");
      } catch (error) {
        setDetectorSupported(false);
        setStatus(error instanceof DOMException && error.name === "NotAllowedError" ? "Camera permission was denied. You can enter an ISBN manually below." : "BookStats could not start barcode scanning in this browser. You can enter an ISBN manually below.");
      }
    }
    void start();
    return () => { runningRef.current = false; controlsRef.current?.stop(); if (timer) window.clearTimeout(timer); streamRef.current?.getTracks().forEach((track) => track.stop()); };
  }, [onDetected]);

  const normalized = normalizeIsbn(manual);
  const validManual = isBookIsbn(normalized);

  return <div className="modal-backdrop scanner-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className="barcode-scanner-modal"><div className="form-header"><div><p className="eyebrow">Add book</p><h2>Scan ISBN barcode</h2></div><button className="icon-button" onClick={onClose} aria-label="Close"><X size={20} /></button></div><div className={`scanner-viewport ${cameraReady ? "ready" : ""}`}><video ref={videoRef} playsInline muted /><div className="scanner-guide"><span /><ScanLine size={34} /></div>{!cameraReady && <div className="scanner-camera-placeholder"><Camera size={38} /><span>{detectorSupported ? "Opening camera…" : "Camera scanner unavailable"}</span></div>}</div><p className="scanner-status">{status}</p><div className="scanner-manual"><div><Keyboard size={17} /><div><strong>Enter ISBN instead</strong><span>Useful if the browser does not support barcode detection or the label is damaged.</span></div></div><div className="scanner-manual-row"><input inputMode="numeric" value={manual} onChange={(event) => setManual(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter" && validManual) onDetected(normalized); }} placeholder="978…" /><button className="button primary compact" disabled={!validManual} onClick={() => onDetected(normalized)}><CheckCircle2 size={15} />Use ISBN</button></div></div></section></div>;
}

function isBookIsbn(value: string): boolean {
  if (value.length === 13) return /^97[89]\d{10}$/.test(value);
  if (value.length !== 10 || !/^\d{9}[\dX]$/.test(value)) return false;
  let sum = 0;
  for (let i = 0; i < 10; i += 1) sum += (10 - i) * (value[i] === "X" ? 10 : Number(value[i]));
  return sum % 11 === 0;
}
