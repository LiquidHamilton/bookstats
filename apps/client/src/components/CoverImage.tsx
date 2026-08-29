import { useEffect, useMemo, useState, type ImgHTMLAttributes, type ReactNode } from "react";
import { coverSources, type CoverDisplayRecord } from "../data/covers";

interface Props extends Omit<ImgHTMLAttributes<HTMLImageElement>, "src"> {
  book: CoverDisplayRecord;
  fallback?: ReactNode;
}

/**
 * Displays the best available copy of a selected cover and automatically falls back when a
 * local cache, archived BookStats asset, or legacy source URL cannot be loaded.
 */
export function CoverImage({ book, fallback = null, onError, ...imageProps }: Props) {
  const sources = useMemo(() => coverSources(book), [
    book.cachedCoverDataUrl,
    book.coverAssetId,
    book.coverAssetToken,
    book.coverUrl,
    book.coverSourceUrl,
    book.updatedAt
  ]);
  const signature = sources.join("\u0000");
  const [sourceIndex, setSourceIndex] = useState(0);

  useEffect(() => { setSourceIndex(0); }, [signature]);

  const source = sources[sourceIndex];
  if (!source) return <>{fallback}</>;

  return <img
    {...imageProps}
    key={`${sourceIndex}:${source}`}
    src={source}
    onError={(event) => {
      onError?.(event);
      setSourceIndex((index) => index + 1);
    }}
  />;
}
