import { useEffect, useRef } from "react";
import * as echarts from "echarts";
import type { EChartsOption } from "echarts";

interface Props { option: EChartsOption; }

export function StatChart({ option }: Props) {
  const target = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!target.current) return;
    const chart = echarts.init(target.current, undefined, { renderer: "canvas" });
    chart.setOption(option);
    const observer = new ResizeObserver(() => chart.resize());
    observer.observe(target.current);
    return () => { observer.disconnect(); chart.dispose(); };
  }, [option]);

  return <div className="stat-chart" ref={target} />;
}
