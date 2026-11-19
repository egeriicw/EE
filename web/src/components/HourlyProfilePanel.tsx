import { useMemo, useState } from "react";
import { analysisApi, type DatasetSummary, type HourlyProfile } from "../api";
import { rampColor, SERIES_COLORS, valueColumns } from "./chartTheme";

interface Props {
  dataset: DatasetSummary;
}

type View = "3d" | "heatmap";

const W = 720;
const H = 420;
const PAD = 16;

export default function HourlyProfilePanel({ dataset }: Props) {
  const values = valueColumns(dataset);
  const [variable, setVariable] = useState(values[values.length - 1] ?? "");
  const [group, setGroup] = useState(dataset.grouped ? dataset.group_names[0] : "");
  const [theta, setTheta] = useState(45);
  const [view, setView] = useState<View>("3d");
  const [profile, setProfile] = useState<HourlyProfile | null>(null);
  const [error, setError] = useState<string | null>(null);

  const run = async () => {
    setError(null);
    try {
      setProfile(await analysisApi.hourlyProfile(dataset.id, { variable, group: group || null, theta_deg: theta }));
    } catch (e) {
      setError((e as Error).message);
    }
  };

  return (
    <section className="panel">
      <h2>Hourly profile</h2>
      <div className="row">
        <label>
          Variable
          <select value={variable} onChange={(e) => setVariable(e.target.value)}>
            {values.map((c) => (
              <option key={c}>{c}</option>
            ))}
          </select>
        </label>
        {dataset.grouped && (
          <label>
            Group
            <select value={group} onChange={(e) => setGroup(e.target.value)}>
              <option value="">all records</option>
              {dataset.group_names.map((g) => (
                <option key={g}>{g}</option>
              ))}
            </select>
          </label>
        )}
        <label>
          View angle: {theta}°
          <input type="range" min={0} max={90} value={theta} onChange={(e) => setTheta(Number(e.target.value))} />
        </label>
        <label>
          Display
          <select value={view} onChange={(e) => setView(e.target.value as View)}>
            <option value="3d">3D day lines</option>
            <option value="heatmap">Heatmap (day × hour)</option>
          </select>
        </label>
        <button onClick={run} disabled={dataset.interval !== "hourly"}>
          Plot profile
        </button>
      </div>
      {dataset.interval !== "hourly" && <p className="muted">Hourly profiles need hourly data.</p>}
      {error && <p className="error">{error}</p>}
      {profile && (
        <>
          {view === "3d" ? <DayLines profile={profile} /> : <Heatmap profile={profile} />}
          <p className="muted">
            {profile.dates.length} days · {profile.n.toLocaleString()} values · scale to {profile.scale.toLocaleString()}{" "}
            {profile.variable}
          </p>
        </>
      )}
    </section>
  );
}

/** Each day is a polyline in the rotated 3D projection (back to front, colour by date). */
function DayLines({ profile }: { profile: HourlyProfile }) {
  const geometry = useMemo(() => {
    const pts = profile.polylines.flatMap((l) => l.points.filter((p): p is [number, number] => p !== null));
    if (pts.length === 0) return null;
    const xs = pts.map((p) => p[0]);
    const ys = pts.map((p) => p[1]);
    const [minX, maxX] = [Math.min(...xs), Math.max(...xs)];
    const [minY, maxY] = [Math.min(...ys), Math.max(...ys)];
    const sx = (x: number) => PAD + ((x - minX) / (maxX - minX || 1)) * (W - 2 * PAD);
    const sy = (y: number) => H - PAD - ((y - minY) / (maxY - minY || 1)) * (H - 2 * PAD);
    return { sx, sy };
  }, [profile]);

  if (!geometry) return <p className="muted">No values to draw.</p>;
  const n = profile.polylines.length;
  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" role="img" aria-label="Hourly profile, one line per day">
      <rect x={0} y={0} width={W} height={H} fill="none" />
      {profile.polylines.map((line, k) => {
        const d = line.points
          .map((p, i) => (p ? `${i === 0 || line.points[i - 1] === null ? "M" : "L"}${geometry.sx(p[0]).toFixed(1)},${geometry.sy(p[1]).toFixed(1)}` : ""))
          .join(" ");
        const t = n > 1 ? k / (n - 1) : 0;
        return (
          <path
            key={line.date}
            d={d}
            fill="none"
            stroke={SERIES_COLORS[0]}
            strokeOpacity={0.12 + 0.5 * t}
            strokeWidth={1}
          >
            <title>{line.date}</title>
          </path>
        );
      })}
      <text x={PAD} y={H - 2} fontSize={11} fill="currentColor">
        one line per day · hover a line for its date
      </text>
    </svg>
  );
}

/** Day across the x-axis, hour of day down the y-axis, colour = value / scale. */
function Heatmap({ profile }: { profile: HourlyProfile }) {
  const cellW = (W - 2 * PAD) / Math.max(1, profile.dates.length);
  const cellH = (H - 2 * PAD) / 24;
  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" role="img" aria-label="Heatmap of hour by day">
      {profile.values.map((row, k) =>
        row.map((v, h) =>
          v === null ? null : (
            <rect
              key={`${k}-${h}`}
              x={PAD + k * cellW}
              y={PAD + h * cellH}
              width={cellW + 0.3}
              height={cellH + 0.3}
              fill={rampColor(v / profile.scale)}
            >
              <title>{`${profile.dates[k]} ${String(h).padStart(2, "0")}:00 — ${v.toLocaleString(undefined, { maximumFractionDigits: 1 })}`}</title>
            </rect>
          ),
        ),
      )}
      {[0, 6, 12, 18, 23].map((h) => (
        <text key={h} x={2} y={PAD + h * cellH + 9} fontSize={10} fill="currentColor">
          {h}
        </text>
      ))}
    </svg>
  );
}
