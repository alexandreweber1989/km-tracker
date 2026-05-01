import { useState, useEffect, useRef, useMemo } from 'react';
import { MapPin, Flag, Save, Download, Trash2, Loader2, AlertCircle, X, Navigation, ChevronDown, ChevronUp, Sparkles, ArrowDown, DollarSign, Calendar, TrendingUp, History } from 'lucide-react';
import * as XLSX from 'xlsx';

const RATE = 1.14; // R$ por km

const SEED_TRIPS = [
  { id: 's1', date: '20/04/2026', origin: 'Rua Attilio Ceccarelli, 90 - Jardim Rio Pequeno, São Paulo - SP, 05388-040', destination: 'Rua dos Marianos, 349 - Centro, Osasco - SP, 06016-050', km: null, geometry: null },
  { id: 's2', date: '20/04/2026', origin: 'Rua dos Marianos, 349 - Centro, Osasco - SP, 06016-050', destination: 'Avenida Marechal Rondon, 199 - Centro, Osasco - SP, 06093-020', km: null, geometry: null },
  { id: 's3', date: '20/04/2026', origin: 'Avenida Marechal Rondon, 165 - Centro, Osasco - SP, 06093-020', destination: 'Rua Antônio Agú, 833 - Centro, Osasco - SP, 06013-000', km: null, geometry: null },
  { id: 's4', date: '20/04/2026', origin: 'Rua Antônio Agú, 833 - Centro, Osasco - SP, 06013-000', destination: 'Rua Minas Bogasian, 284 - Centro, Osasco - SP, 06013-010', km: null, geometry: null },
  { id: 's5', date: '20/04/2026', origin: 'Rua Minas Bogasian, 284 - Centro, Osasco - SP, 06013-010', destination: 'Rua Dona Primitiva Vianco, 589 - Centro, Osasco - SP, 06010-004', km: null, geometry: null },
  { id: 's6', date: '20/04/2026', origin: 'Rua Dona Primitiva Vianco, 589 - Centro, Osasco - SP, 06010-004', destination: 'Avenida João Batista, 11 - Centro, Osasco - SP, 06097-100', km: null, geometry: null },
  { id: 's7', date: '20/04/2026', origin: 'Avenida João Batista, 11 - Centro, Osasco - SP, 06097-100', destination: 'Avenida João Batista, 253 - Centro, Osasco - SP, 06090-100', km: null, geometry: null },
  { id: 's8', date: '20/04/2026', origin: 'Avenida João Batista, 253 - Centro, Osasco - SP, 06090-100', destination: 'Rua Fiorino Beltrano, 195 - Centro, Osasco - SP, 06097-040', km: null, geometry: null },
  { id: 's9', date: '20/04/2026', origin: 'Rua Fiorino Beltrano, 195 - Centro, Osasco - SP, 06097-040', destination: 'Rua Attilio Ceccarelli, 90 - Jardim Rio Pequeno, São Paulo - SP, 05388-040', km: null, geometry: null },
];

const STATE_MAP = {
  'Acre':'AC','Alagoas':'AL','Amapá':'AP','Amazonas':'AM','Bahia':'BA','Ceará':'CE',
  'Distrito Federal':'DF','Espírito Santo':'ES','Goiás':'GO','Maranhão':'MA',
  'Mato Grosso':'MT','Mato Grosso do Sul':'MS','Minas Gerais':'MG','Pará':'PA',
  'Paraíba':'PB','Paraná':'PR','Pernambuco':'PE','Piauí':'PI','Rio de Janeiro':'RJ',
  'Rio Grande do Norte':'RN','Rio Grande do Sul':'RS','Rondônia':'RO','Roraima':'RR',
  'Santa Catarina':'SC','São Paulo':'SP','Sergipe':'SE','Tocantins':'TO'
};

const C = {
  red:'#E61A27', redDeep:'#B80F1B', redDark:'#7A0A12', redBright:'#FF2738',
  bg:'#F4F1ED', ink:'#0A0908', inkSoft:'#4A4744', inkFaded:'#8F8B86',
  card:'#FFFFFF', border:'#E8E0D6', borderDark:'#C9BDAE', white:'#FFFFFF', black:'#080606',
  green:'#16A34A',
};

const MONTH_NAMES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

function formatBrazilianAddress(data) {
  const a = data.address || {};
  const street = a.road || a.pedestrian || a.path || '';
  const number = a.house_number || '';
  const neighborhood = a.suburb || a.neighbourhood || a.quarter || a.city_district || '';
  const city = a.city || a.town || a.village || a.municipality || '';
  const stateAbbr = STATE_MAP[a.state || ''] || a.state || '';
  const postcode = a.postcode || '';
  let parts = [];
  if (street) {
    let s = number ? `${street}, ${number}` : street;
    if (neighborhood) s += ` - ${neighborhood}`;
    parts.push(s);
  } else if (neighborhood) parts.push(neighborhood);
  let loc = city;
  if (stateAbbr) loc += (loc ? ' - ' : '') + stateAbbr;
  if (postcode) loc += (loc ? ', ' : '') + postcode;
  if (loc) parts.push(loc);
  return parts.length > 0 ? parts.join(', ') : (data.display_name || '');
}

function haversineKm(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

const formatBRL = (n) => 'R$ ' + n.toFixed(2).replace('.', ',');

const STORAGE_KEY = 'km_trips_v1';
function loadTrips() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (_) {}
  return null;
}
function saveTripsToStorage(trips) {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(trips)); } catch (_) {}
}

// ── MINI MAP ──────────────────────────────────────────────────────────────────
function MiniMap({ geometry, fM }) {
  if (!geometry || geometry.length < 2) {
    return (
      <div style={{
        width: '100%', height: 88, borderRadius: 12,
        background: `repeating-linear-gradient(45deg, ${C.bg}, ${C.bg} 8px, #ECE5DA 8px, #ECE5DA 16px)`,
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        border: `1px dashed ${C.borderDark}`, marginTop: 10
      }}>
        <span style={{ fontFamily: fM, fontSize: 10, color: C.inkFaded, letterSpacing: '0.14em', textTransform: 'uppercase' }}>
          ◊ rota não disponível ◊
        </span>
      </div>
    );
  }

  let minLng = Infinity, maxLng = -Infinity, minLat = Infinity, maxLat = -Infinity;
  for (const [lng, lat] of geometry) {
    if (lng < minLng) minLng = lng;
    if (lng > maxLng) maxLng = lng;
    if (lat < minLat) minLat = lat;
    if (lat > maxLat) maxLat = lat;
  }
  const padX = ((maxLng - minLng) || 0.001) * 0.15;
  const padY = ((maxLat - minLat) || 0.001) * 0.20;
  minLng -= padX; maxLng += padX;
  minLat -= padY; maxLat += padY;

  const W = 320, H = 88;
  const dataAspect = (maxLng - minLng) / (maxLat - minLat);
  const viewAspect = W / H;
  if (dataAspect > viewAspect) {
    const newLatRange = (maxLng - minLng) / viewAspect;
    const midLat = (minLat + maxLat) / 2;
    minLat = midLat - newLatRange / 2;
    maxLat = midLat + newLatRange / 2;
  } else {
    const newLngRange = (maxLat - minLat) * viewAspect;
    const midLng = (minLng + maxLng) / 2;
    minLng = midLng - newLngRange / 2;
    maxLng = midLng + newLngRange / 2;
  }

  const project = ([lng, lat]) => {
    const x = ((lng - minLng) / (maxLng - minLng)) * W;
    const y = H - ((lat - minLat) / (maxLat - minLat)) * H;
    return [x, y];
  };

  const points = geometry.map(project);
  const pathD = points.map(([x, y], i) => `${i === 0 ? 'M' : 'L'}${x.toFixed(1)},${y.toFixed(1)}`).join(' ');
  const [sx, sy] = points[0];
  const [ex, ey] = points[points.length - 1];

  return (
    <div style={{ marginTop: 10, position: 'relative', borderRadius: 12, overflow: 'hidden', border: `1px solid ${C.border}` }}>
      <svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="xMidYMid meet" style={{
        width: '100%', height: 88,
        background: `linear-gradient(135deg, #F8F4ED 0%, #EDE5D8 100%)`,
        display: 'block'
      }}>
        {/* Subtle grid */}
        <defs>
          <pattern id="mapgrid" width="20" height="20" patternUnits="userSpaceOnUse">
            <path d="M 20 0 L 0 0 0 20" fill="none" stroke="rgba(10,9,8,0.05)" strokeWidth="0.5" />
          </pattern>
        </defs>
        <rect width={W} height={H} fill="url(#mapgrid)" />

        {/* Route shadow + line */}
        <path d={pathD} fill="none" stroke="rgba(230,26,39,0.25)" strokeWidth="6" strokeLinecap="round" strokeLinejoin="round" />
        <path d={pathD} fill="none" stroke={C.red} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />

        {/* Start marker */}
        <circle cx={sx} cy={sy} r="6" fill="white" stroke={C.red} strokeWidth="2.5" />
        <circle cx={sx} cy={sy} r="2.5" fill={C.red} />

        {/* End marker (flag) */}
        <rect x={ex - 5} y={ey - 5} width="10" height="10" fill={C.black} stroke="white" strokeWidth="2" rx="1.5" />
      </svg>
    </div>
  );
}

// ── DASHBOARD ─────────────────────────────────────────────────────────────────
function Dashboard({ trips, fD, fM, fB }) {
  const stats = useMemo(() => {
    const totalKm = trips.reduce((s, t) => s + (t.km || 0), 0);
    const measured = trips.filter(t => t.km != null);

    const byDay = {};
    for (const t of measured) {
      if (!byDay[t.date]) byDay[t.date] = { km: 0, count: 0 };
      byDay[t.date].km += t.km;
      byDay[t.date].count += 1;
    }
    const days = Object.entries(byDay).map(([date, d]) => ({ date, ...d }))
      .sort((a, b) => {
        const p = s => { const [d,m,y] = s.split('/'); return new Date(y,m-1,d); };
        return p(b.date) - p(a.date);
      });

    const byMonth = {};
    for (const t of measured) {
      const [d, m, y] = t.date.split('/');
      const key = `${m}/${y}`;
      if (!byMonth[key]) byMonth[key] = { km: 0, count: 0 };
      byMonth[key].km += t.km;
      byMonth[key].count += 1;
    }
    const months = Object.entries(byMonth).map(([key, d]) => ({ key, ...d }))
      .sort((a, b) => {
        const [ma, ya] = a.key.split('/');
        const [mb, yb] = b.key.split('/');
        return new Date(yb, mb-1) - new Date(ya, ma-1);
      });

    const maxDayKm = Math.max(...days.map(d => d.km), 0.01);

    return { totalKm, measuredCount: measured.length, days, months, maxDayKm };
  }, [trips]);

  const totalEarning = stats.totalKm * RATE;

  const cardBase = {
    backgroundColor: C.card, borderRadius: 22, padding: 22, marginBottom: 14,
    boxShadow: '0 2px 6px rgba(10,9,8,0.04), 0 16px 40px rgba(10,9,8,0.06)',
    border: `1px solid ${C.border}`
  };

  const sectionHeader = (icon, title, badge) => (
    <div style={{ display:'flex', alignItems:'center', gap:12, marginBottom:16 }}>
      <div style={{ width:4, height:24, borderRadius:2, background:`linear-gradient(180deg, ${C.red}, ${C.redDeep})` }} />
      <div style={{ display:'flex', alignItems:'center', gap:8 }}>
        {icon}
        <h3 style={{ fontFamily:fD, fontSize:22, color:C.ink, letterSpacing:'-0.03em', fontWeight:700, margin:0 }}>{title}</h3>
      </div>
      {badge && <span style={{ fontFamily:fM, fontSize:10, color:C.white, background:C.black, padding:'3px 8px', borderRadius:6, letterSpacing:'0.06em' }}>{badge}</span>}
    </div>
  );

  return (
    <div className="km-fade">
      {/* MAIN EARNING CARD */}
      <div style={{
        background: `linear-gradient(135deg, ${C.black} 0%, #1A1414 100%)`,
        borderRadius: 24, padding: '24px 22px', marginBottom: 14,
        position: 'relative', overflow: 'hidden',
        border: `1px solid rgba(230,26,39,0.3)`,
        boxShadow: '0 18px 44px rgba(10,9,8,0.18)'
      }}>
        {/* Glow */}
        <div style={{
          position: 'absolute', top: -50, right: -30, width: 180, height: 180,
          borderRadius: '50%', background: `radial-gradient(circle, rgba(230,26,39,0.45), transparent 70%)`,
          pointerEvents: 'none'
        }} />
        <div style={{
          position: 'absolute', inset: 0, opacity: 0.08,
          backgroundImage: 'linear-gradient(rgba(255,255,255,0.4) 1px, transparent 1px)',
          backgroundSize: '100% 14px', pointerEvents: 'none'
        }} />

        <div style={{ position:'relative' }}>
          <div style={{
            display:'flex', alignItems:'center', gap:8,
            fontFamily:fM, fontSize:10, letterSpacing:'0.26em',
            color:'rgba(255,255,255,0.6)', textTransform:'uppercase', fontWeight:500
          }}>
            <DollarSign size={12} strokeWidth={2.4} />
            Valor a receber
          </div>
          <div style={{ display:'flex', alignItems:'baseline', gap:8, marginTop:10 }}>
            <span style={{ fontFamily:fD, fontSize:22, color:'rgba(255,255,255,0.7)', fontWeight:700 }}>R$</span>
            <span style={{
              fontFamily:fD, fontSize:60, lineHeight:0.9, color:C.white,
              fontWeight:800, letterSpacing:'-0.05em',
              textShadow:'0 0 28px rgba(230,26,39,0.5)'
            }}>
              {totalEarning.toFixed(2).replace('.',',')}
            </span>
          </div>
          <div style={{ marginTop:14, display:'flex', alignItems:'center', gap:10, flexWrap:'wrap' }}>
            <span style={{
              fontFamily:fM, fontSize:11, color:'rgba(255,255,255,0.7)',
              background:'rgba(255,255,255,0.08)', padding:'5px 10px', borderRadius:100,
              border:'1px solid rgba(255,255,255,0.12)', letterSpacing:'0.06em'
            }}>
              {stats.totalKm.toFixed(2)} km
            </span>
            <span style={{
              fontFamily:fM, fontSize:11, color:'rgba(255,255,255,0.5)', letterSpacing:'0.06em'
            }}>
              × R$ 1,14/km
            </span>
          </div>
          <div style={{
            marginTop:10, fontFamily:fM, fontSize:10, color:'rgba(255,255,255,0.4)',
            letterSpacing:'0.08em'
          }}>
            {stats.measuredCount} de {trips.length} viagens medidas
          </div>
        </div>
      </div>

      {/* POR DIA */}
      <div style={cardBase}>
        {sectionHeader(
          <Calendar size={18} color={C.ink} strokeWidth={2.2} />,
          'Por dia',
          String(stats.days.length).padStart(2,'0')
        )}

        {stats.days.length === 0 ? (
          <p style={{ fontSize:12.5, color:C.inkFaded, fontStyle:'italic', textAlign:'center', padding:'12px 0', fontFamily:fD, margin:0 }}>
            Sem viagens medidas ainda.
          </p>
        ) : (
          <div>
            {stats.days.map((d, idx) => {
              const earning = d.km * RATE;
              const barPct = (d.km / stats.maxDayKm) * 100;
              return (
                <div key={d.date} style={{
                  paddingTop: idx === 0 ? 0 : 14,
                  paddingBottom: 4,
                  borderTop: idx === 0 ? 'none' : `1px solid ${C.border}`
                }}>
                  <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:8 }}>
                    <div>
                      <div style={{ fontFamily:fM, fontSize:12, fontWeight:700, color:C.ink, letterSpacing:'0.1em' }}>
                        {d.date}
                      </div>
                      <div style={{ fontFamily:fM, fontSize:10, color:C.inkFaded, marginTop:3, letterSpacing:'0.06em' }}>
                        {d.count} {d.count === 1 ? 'viagem' : 'viagens'}
                      </div>
                    </div>
                    <div style={{ textAlign:'right' }}>
                      <div style={{ fontFamily:fM, fontSize:14, fontWeight:700, color:C.ink, letterSpacing:'-0.01em' }}>
                        {d.km.toFixed(2)} km
                      </div>
                      <div style={{ fontFamily:fD, fontSize:13, fontWeight:700, color:C.red, marginTop:2, fontStyle:'italic' }}>
                        {formatBRL(earning)}
                      </div>
                    </div>
                  </div>
                  <div style={{ height:5, background:C.bg, borderRadius:3, overflow:'hidden' }}>
                    <div style={{
                      width:`${barPct}%`, height:'100%',
                      background:`linear-gradient(90deg, ${C.red}, ${C.redDeep})`,
                      borderRadius:3, transition:'width 0.4s ease'
                    }} />
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* POR MÊS */}
      <div style={{ ...cardBase, marginBottom: 0 }}>
        {sectionHeader(
          <TrendingUp size={18} color={C.ink} strokeWidth={2.2} />,
          'Por mês',
          String(stats.months.length).padStart(2,'0')
        )}

        {stats.months.length === 0 ? (
          <p style={{ fontSize:12.5, color:C.inkFaded, fontStyle:'italic', textAlign:'center', padding:'12px 0', fontFamily:fD, margin:0 }}>
            Sem dados mensais ainda.
          </p>
        ) : (
          <div>
            {stats.months.map((m, idx) => {
              const [mm, yy] = m.key.split('/');
              const label = `${MONTH_NAMES[parseInt(mm)-1]} · ${yy}`;
              const earning = m.km * RATE;
              return (
                <div key={m.key} style={{
                  paddingTop: idx === 0 ? 0 : 14,
                  paddingBottom: idx === stats.months.length - 1 ? 0 : 0,
                  borderTop: idx === 0 ? 'none' : `1px solid ${C.border}`,
                  display:'flex', alignItems:'center', justifyContent:'space-between'
                }}>
                  <div>
                    <div style={{ fontFamily:fD, fontSize:18, fontWeight:700, color:C.ink, letterSpacing:'-0.02em' }}>
                      {label}
                    </div>
                    <div style={{ fontFamily:fM, fontSize:10.5, color:C.inkFaded, marginTop:3, letterSpacing:'0.08em' }}>
                      {m.count} {m.count === 1 ? 'viagem' : 'viagens'} · {m.km.toFixed(2)} km
                    </div>
                  </div>
                  <div style={{
                    fontFamily:fD, fontSize:18, fontWeight:700, color:C.white,
                    background:`linear-gradient(135deg, ${C.red}, ${C.redDeep})`,
                    padding:'8px 14px', borderRadius:100, letterSpacing:'-0.01em',
                    boxShadow:'0 6px 16px rgba(230,26,39,0.28)', fontStyle:'italic'
                  }}>
                    {formatBRL(earning)}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

// ── MAIN APP ──────────────────────────────────────────────────────────────────
export default function KmTracker() {
  const [origin, setOrigin] = useState({ address: '', lat: null, lng: null });
  const [destination, setDestination] = useState({ address: '', lat: null, lng: null });
  const [distance, setDistance] = useState(null);
  const [distanceLabel, setDistanceLabel] = useState('');
  const [routeGeometry, setRouteGeometry] = useState(null);
  const [trips, setTrips] = useState([]);
  const [loading, setLoading] = useState({ origin: false, destination: false, distance: false });
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [historyOpen, setHistoryOpen] = useState(true);
  const [activeTab, setActiveTab] = useState('history'); // 'history' | 'dashboard'
  const [time, setTime] = useState('');
  const [originSuggestions, setOriginSuggestions] = useState([]);
  const [destSuggestions, setDestSuggestions] = useState([]);
  const initialized = useRef(false);
  const originTyping = useRef(false);
  const destTyping = useRef(false);

  useEffect(() => {
    const update = () => {
      const d = new Date();
      setTime(`${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`);
    };
    update();
    const id = setInterval(update, 1000);
    return () => clearInterval(id);
  }, []);

  useEffect(() => {
    const loaded = loadTrips();
    if (loaded && Array.isArray(loaded)) setTrips(loaded);
    else { setTrips(SEED_TRIPS); saveTripsToStorage(SEED_TRIPS); }
  }, []);

  useEffect(() => {
    if (initialized.current) return;
    initialized.current = true;
    const fonts = document.createElement('link');
    fonts.href = 'https://fonts.googleapis.com/css2?family=Syne:wght@500;600;700;800&family=Inter+Tight:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap';
    fonts.rel = 'stylesheet';
    document.head.appendChild(fonts);
    const style = document.createElement('style');
    style.textContent = `
      * { box-sizing: border-box; }
      body { margin: 0; padding: 0; }
      @keyframes bubbleRise {
        0% { transform: translateY(0) scale(0.6); opacity: 0; }
        15% { opacity: 0.7; } 85% { opacity: 0.5; }
        100% { transform: translateY(-260px) scale(1.1); opacity: 0; }
      }
      @keyframes fadeUp {
        from { opacity: 0; transform: translateY(8px); }
        to { opacity: 1; transform: translateY(0); }
      }
      @keyframes shimmer {
        0% { background-position: -200% center; }
        100% { background-position: 200% center; }
      }
      @keyframes blink {
        0%,49% { opacity: 1; } 50%,100% { opacity: 0.3; }
      }
      @keyframes spin {
        from { transform: rotate(0); } to { transform: rotate(360deg); }
      }
      .km-bubble {
        position: absolute; border-radius: 50%;
        background: radial-gradient(circle at 35% 35%, rgba(255,255,255,0.95), rgba(255,255,255,0.2) 70%);
        animation: bubbleRise linear infinite; pointer-events: none;
      }
      .km-press:active { transform: scale(0.97); transition: transform 0.08s ease; }
      .km-fade { animation: fadeUp 0.35s ease-out; }
      .km-shimmer {
        background: linear-gradient(90deg, #FFF 0%, #FFE5E7 50%, #FFF 100%);
        background-size: 200% 100%;
        animation: shimmer 3.5s linear infinite;
        -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent;
      }
      .km-blink { animation: blink 1.2s step-end infinite; }
      .km-spin { animation: spin 1s linear infinite; }
      textarea:focus, input:focus { border-color: #E61A27 !important; outline: none; }
      button { cursor: pointer; border: none; }
    `;
    document.head.appendChild(style);
  }, []);

  useEffect(() => {
    if (trips.length > 0 || loadTrips()) saveTripsToStorage(trips);
  }, [trips]);

  // Auto-calc distance + geometry
  useEffect(() => {
    if (origin.lat == null || destination.lat == null) {
      setDistance(null); setDistanceLabel(''); setRouteGeometry(null); return;
    }
    let cancelled = false;
    (async () => {
      setLoading(s => ({ ...s, distance: true }));
      try {
        const r = await fetch(`https://router.project-osrm.org/route/v1/driving/${origin.lng},${origin.lat};${destination.lng},${destination.lat}?overview=simplified&geometries=geojson`);
        const data = await r.json();
        if (cancelled) return;
        if (data.routes?.length > 0) {
          const route = data.routes[0];
          setDistance(route.distance / 1000);
          setDistanceLabel('rota de carro');
          setRouteGeometry(route.geometry?.coordinates || null);
        } else throw new Error();
      } catch {
        if (cancelled) return;
        setDistance(haversineKm(origin.lat, origin.lng, destination.lat, destination.lng));
        setDistanceLabel('linha reta · rota indisponível');
        setRouteGeometry([[origin.lng, origin.lat], [destination.lng, destination.lat]]);
      } finally {
        if (!cancelled) setLoading(s => ({ ...s, distance: false }));
      }
    })();
    return () => { cancelled = true; };
  }, [origin.lat, origin.lng, destination.lat, destination.lng]);

  // Autocomplete origin
  useEffect(() => {
    if (!originTyping.current) return;
    const q = origin.address.trim();
    if (q.length < 3) { setOriginSuggestions([]); return; }
    const timer = setTimeout(async () => {
      try {
        const r = await fetch(`https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(q)}&format=json&addressdetails=1&limit=5&accept-language=pt-BR&countrycodes=br`);
        const data = await r.json();
        if (originTyping.current) setOriginSuggestions(data.map(d => ({ display: formatBrazilianAddress(d), lat: +d.lat, lng: +d.lon })));
      } catch { setOriginSuggestions([]); }
    }, 400);
    return () => clearTimeout(timer);
  }, [origin.address]);

  // Autocomplete destination
  useEffect(() => {
    if (!destTyping.current) return;
    const q = destination.address.trim();
    if (q.length < 3) { setDestSuggestions([]); return; }
    const timer = setTimeout(async () => {
      try {
        const r = await fetch(`https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(q)}&format=json&addressdetails=1&limit=5&accept-language=pt-BR&countrycodes=br`);
        const data = await r.json();
        if (destTyping.current) setDestSuggestions(data.map(d => ({ display: formatBrazilianAddress(d), lat: +d.lat, lng: +d.lon })));
      } catch { setDestSuggestions([]); }
    }, 400);
    return () => clearTimeout(timer);
  }, [destination.address]);

  function selectSuggestion(which, suggestion) {
    const loc = { address: suggestion.display, lat: suggestion.lat, lng: suggestion.lng };
    if (which === 'origin') {
      originTyping.current = false;
      setOrigin(loc);
      setOriginSuggestions([]);
    } else {
      destTyping.current = false;
      setDestination(loc);
      setDestSuggestions([]);
    }
  }

  async function capture(which) {
    setError(null); setLoading(s => ({ ...s, [which]: true }));
    try {
      if (!navigator.geolocation) throw new Error('Geolocalização não disponível neste navegador');
      const pos = await new Promise((res, rej) => navigator.geolocation.getCurrentPosition(
        p => res({ lat: p.coords.latitude, lng: p.coords.longitude }),
        e => rej(new Error(
          e.code === 1 ? 'Permissão de localização negada. Habilite o GPS para este site.' :
          e.code === 2 ? 'Localização indisponível. Verifique se o GPS está ligado.' :
          e.code === 3 ? 'Tempo esgotado. Tente novamente.' : 'Erro ao obter localização'
        )),
        { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 }
      ));
      const r = await fetch(`https://nominatim.openstreetmap.org/reverse?lat=${pos.lat}&lon=${pos.lng}&format=json&accept-language=pt-BR&addressdetails=1`);
      const data = await r.json();
      const address = formatBrazilianAddress(data);
      const loc = { address, lat: pos.lat, lng: pos.lng };
      if (which === 'origin') {
        setOrigin(loc); setOriginSuggestions([]);
      } else {
        setDestination(loc); setDestSuggestions([]);
      }
    } catch (e) { setError(e.message); }
    finally { setLoading(s => ({ ...s, [which]: false })); }
  }

  function clearLocation(which) {
    if (which === 'origin') {
      setOrigin({ address: '', lat: null, lng: null });
      setOriginSuggestions([]);
    } else {
      setDestination({ address: '', lat: null, lng: null });
      setDestSuggestions([]);
    }
  }

  function handleSave() {
    if (!origin.address.trim() || !destination.address.trim()) { setError('Preencha origem e destino'); return; }
    const d = new Date();
    const newTrip = {
      id: 't' + Date.now(),
      date: `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`,
      origin: origin.address.trim(),
      destination: destination.address.trim(),
      km: distance != null ? Number(distance.toFixed(2)) : null,
      kmLabel: distanceLabel,
      geometry: routeGeometry,
    };
    setTrips(prev => [newTrip, ...prev]);
    setOrigin({ address:'', lat:null, lng:null });
    setDestination({ address:'', lat:null, lng:null });
    setOriginSuggestions([]); setDestSuggestions([]);
    setDistance(null); setDistanceLabel(''); setRouteGeometry(null); setError(null);
    setSuccess('Viagem registrada'); setTimeout(() => setSuccess(null), 2200);
  }

  function deleteTrip(id) {
    if (window.confirm('Remover esta viagem?')) setTrips(prev => prev.filter(t => t.id !== id));
  }

  function exportToExcel() {
    if (!trips.length) { setError('Nenhuma viagem para exportar'); return; }
    const sorted = [...trips].sort((a, b) => {
      const parse = s => { const [d,m,y] = s.split('/'); return new Date(y,m-1,d); };
      return parse(a.date) - parse(b.date);
    });
    const data = [['Data','Origem','Destino'], ...sorted.map(t => [t.date, t.origin, t.destination])];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch:12 },{ wch:60 },{ wch:60 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'kmadicional');
    XLSX.writeFile(wb, 'lançamento_de_Km.xlsx');
    setSuccess('Planilha exportada'); setTimeout(() => setSuccess(null), 2200);
  }

  const totalKm = trips.reduce((s, t) => s + (t.km || 0), 0);
  const totalEarning = totalKm * RATE;
  const tripsWithKm = trips.filter(t => t.km != null).length;
  const canSave = origin.address.trim() && destination.address.trim();

  const fD = "'Syne', system-ui, sans-serif";
  const fM = "'JetBrains Mono', monospace";
  const fB = "'Inter Tight', system-ui, sans-serif";

  const bubbles = [
    {left:'8%',size:14,delay:0,dur:7},{left:'22%',size:8,delay:2.2,dur:6},
    {left:'38%',size:18,delay:4.5,dur:8},{left:'55%',size:10,delay:1.1,dur:6.5},
    {left:'70%',size:16,delay:3.4,dur:7.5},{left:'85%',size:9,delay:5.7,dur:6},
    {left:'92%',size:12,delay:0.6,dur:7.2},{left:'15%',size:6,delay:3.9,dur:5.5},
    {left:'48%',size:7,delay:5.2,dur:6.8},
  ];

  const s = {
    card: { backgroundColor:C.card, borderRadius:26, padding:24, marginBottom:18,
      boxShadow:'0 2px 6px rgba(10,9,8,0.04), 0 16px 40px rgba(10,9,8,0.06)', border:`1px solid ${C.border}` },
    redBar: { width:4, height:28, borderRadius:2, background:`linear-gradient(180deg, ${C.red}, ${C.redDeep})` },
    label: { display:'flex', alignItems:'center', gap:8, fontFamily:fM, fontSize:10.5,
      fontWeight:500, color:C.ink, letterSpacing:'0.18em', textTransform:'uppercase' },
    textarea: { width:'100%', padding:'12px 14px', fontSize:14, fontFamily:fB, color:C.ink,
      backgroundColor:C.bg, border:`1.5px solid ${C.border}`, borderRadius:14,
      resize:'none', lineHeight:1.45, fontWeight:500 },
    btnRed: { width:'100%', marginTop:9, padding:'13px 14px', borderRadius:14,
      background:`linear-gradient(135deg, ${C.red} 0%, ${C.redDeep} 100%)`, color:C.white,
      fontWeight:600, fontSize:13, fontFamily:fD, display:'flex', alignItems:'center',
      justifyContent:'center', gap:8, letterSpacing:'0.04em',
      boxShadow:'0 6px 18px rgba(230,26,39,0.32)' },
    btnBlack: { width:'100%', marginTop:9, padding:'13px 14px', borderRadius:14,
      background:C.black, color:C.white, fontWeight:600, fontSize:13, fontFamily:fD,
      display:'flex', alignItems:'center', justifyContent:'center', gap:8,
      letterSpacing:'0.04em', boxShadow:'0 6px 18px rgba(0,0,0,0.22)' },
  };

  const tabBtn = (active) => ({
    flex: 1, padding: '11px 14px', borderRadius: 100,
    background: active ? C.black : 'transparent',
    color: active ? C.white : C.inkSoft,
    fontFamily: fD, fontSize: 13, fontWeight: 600, letterSpacing: '-0.01em',
    transition: 'all 0.2s ease',
    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 7,
    boxShadow: active ? '0 4px 12px rgba(0,0,0,0.15)' : 'none'
  });

  return (
    <div style={{ backgroundColor:C.bg, fontFamily:fB, color:C.ink, minHeight:'100vh' }}>

      {/* HERO */}
      <div style={{ position:'relative', background:`linear-gradient(160deg, ${C.red} 0%, ${C.redDeep} 55%, ${C.redDark} 100%)`, paddingBottom:64, overflow:'hidden' }}>
        <div style={{ position:'absolute', inset:0, overflow:'hidden' }}>
          {bubbles.map((b,i) => <span key={i} className="km-bubble" style={{ left:b.left, bottom:-20, width:b.size, height:b.size, animationDuration:`${b.dur}s`, animationDelay:`${b.delay}s` }} />)}
        </div>
        <div style={{ position:'absolute', inset:0, opacity:0.08, backgroundImage:'linear-gradient(rgba(255,255,255,0.4) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.4) 1px, transparent 1px)', backgroundSize:'32px 32px', pointerEvents:'none' }} />
        <div style={{ position:'absolute', top:-60, right:-40, width:220, height:220, borderRadius:'50%', background:`radial-gradient(circle, #FF475766, transparent 70%)`, pointerEvents:'none' }} />

        <div style={{ maxWidth:448, margin:'0 auto', padding:'20px 24px 0', position:'relative', zIndex:2 }}>
          <div style={{ display:'flex', justifyContent:'space-between', fontFamily:fM, fontSize:10, color:'rgba(255,255,255,0.7)', letterSpacing:'0.08em' }}>
            <div style={{ display:'flex', alignItems:'center', gap:6 }}>
              <span className="km-blink" style={{ width:6, height:6, borderRadius:3, background:'#4ADE80', boxShadow:'0 0 6px #4ADE80', display:'inline-block' }} />
              SYS · ONLINE
            </div>
            <span>{time}</span>
          </div>

          <div style={{ marginTop:28, display:'flex', alignItems:'flex-start', justifyContent:'space-between' }}>
            <div>
              <div style={{ fontFamily:fM, fontSize:10, letterSpacing:'0.3em', color:'rgba(255,255,255,0.7)', fontWeight:500 }}>[ v2·0 ] · DRIVE LOG</div>
              <h1 style={{ fontFamily:fD, fontSize:52, lineHeight:0.88, color:C.white, marginTop:8, letterSpacing:'-0.04em', fontWeight:800, textShadow:'0 2px 24px rgba(0,0,0,0.18)', margin:'8px 0 0' }}>
                Controle<br /><span style={{ fontStyle:'italic', fontWeight:700 }}>de </span>KM
              </h1>
            </div>
            <div style={{ width:56, height:56, borderRadius:18, background:'rgba(255,255,255,0.12)', backdropFilter:'blur(12px)', border:'1px solid rgba(255,255,255,0.28)', display:'flex', alignItems:'center', justifyContent:'center', flexShrink:0 }}>
              <Navigation size={24} color={C.white} strokeWidth={2.2} />
            </div>
          </div>

          <div style={{ marginTop:30 }}>
            <div style={{ fontFamily:fM, fontSize:10, letterSpacing:'0.28em', color:'rgba(255,255,255,0.65)', fontWeight:500, display:'flex', alignItems:'center', gap:8 }}>
              <span style={{ width:18, height:1, background:'rgba(255,255,255,0.4)', display:'inline-block' }} />
              Total · {totalKm.toFixed(1)} km
            </div>
            <div style={{ display:'flex', alignItems:'baseline', gap:10, marginTop:6 }}>
              <span style={{ fontFamily:fD, fontSize:22, color:'rgba(255,255,255,0.7)', fontWeight:700 }}>R$</span>
              <div className="km-shimmer" style={{ fontFamily:fD, fontSize:64, lineHeight:0.88, letterSpacing:'-0.05em', fontWeight:800 }}>
                {totalEarning.toFixed(2).replace('.',',')}
              </div>
            </div>
            <div style={{ marginTop:12, display:'inline-flex', alignItems:'center', gap:10, padding:'6px 14px', borderRadius:100, background:'rgba(0,0,0,0.22)', border:'1px solid rgba(255,255,255,0.12)' }}>
              <span className="km-blink" style={{ width:6, height:6, borderRadius:3, background:'#4ADE80', boxShadow:'0 0 8px #4ADE80', display:'inline-block' }} />
              <span style={{ fontFamily:fM, fontSize:10.5, color:C.white, fontWeight:500, letterSpacing:'0.1em' }}>
                {String(trips.length).padStart(3,'0')} trips · {String(tripsWithKm).padStart(3,'0')} medidas
              </span>
            </div>
          </div>
        </div>

        <svg viewBox="0 0 400 60" preserveAspectRatio="none" style={{ position:'absolute', bottom:-1, left:0, width:'100%', height:64, display:'block' }}>
          <path d="M0,40 C60,10 120,55 200,30 C280,5 340,50 400,25 L400,60 L0,60 Z" fill={C.bg} />
          <path d="M0,40 C60,10 120,55 200,30 C280,5 340,50 400,25" fill="none" stroke="rgba(255,255,255,0.35)" strokeWidth="1.5" />
        </svg>
      </div>

      {/* CONTENT */}
      <div style={{ maxWidth:448, margin:'-24px auto 0', padding:'0 20px 48px', position:'relative', zIndex:3 }}>

        {error && (
          <div className="km-fade" style={{ backgroundColor:'#FFF1F2', border:`1.5px solid ${C.red}`, borderRadius:16, padding:'12px 14px', display:'flex', gap:10, alignItems:'flex-start', marginBottom:14, boxShadow:'0 4px 14px rgba(230,26,39,0.12)' }}>
            <AlertCircle size={20} color={C.redDeep} style={{ marginTop:1, flexShrink:0 }} strokeWidth={2.4} />
            <p style={{ fontSize:13, color:C.redDark, flex:1, lineHeight:1.4, fontWeight:500, margin:0 }}>{error}</p>
            <button onClick={() => setError(null)} style={{ color:C.redDeep, padding:2, background:'none' }}><X size={16} /></button>
          </div>
        )}
        {success && (
          <div className="km-fade" style={{ background:`linear-gradient(135deg, ${C.red}, ${C.redDeep})`, borderRadius:16, padding:'12px 16px', display:'flex', gap:10, alignItems:'center', marginBottom:14, boxShadow:'0 8px 24px rgba(230,26,39,0.32)' }}>
            <Sparkles size={16} color={C.white} strokeWidth={2.4} />
            <p style={{ fontSize:14, color:C.white, fontWeight:600, fontFamily:fD, letterSpacing:'-0.01em', margin:0 }}>{success}</p>
          </div>
        )}

        {/* NEW TRIP CARD */}
        <section style={s.card}>
          <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:18 }}>
            <div style={{ display:'flex', alignItems:'center', gap:12 }}>
              <div style={s.redBar} />
              <h2 style={{ fontFamily:fD, fontSize:26, color:C.ink, letterSpacing:'-0.03em', fontWeight:700, margin:0 }}>Nova viagem</h2>
            </div>
            <span style={{ fontFamily:fM, fontSize:9.5, color:C.inkFaded, letterSpacing:'0.16em' }}>001/REC</span>
          </div>

          {/* Origin */}
          <div style={{ marginBottom:14 }}>
            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:10 }}>
              <label style={s.label}>
                <span style={{ width:10, height:10, borderRadius:5, background:`radial-gradient(circle at 30% 30%, ${C.redBright}, ${C.redDeep})`, boxShadow:`0 0 0 3px rgba(230,26,39,0.15)`, display:'inline-block' }} />
                · Origem
              </label>
              {origin.address && <button onClick={() => clearLocation('origin')} style={{ fontFamily:fM, fontSize:10, color:C.inkFaded, background:'none', letterSpacing:'0.1em', textTransform:'uppercase' }}>clear</button>}
            </div>
            <div style={{ position:'relative' }}>
              <textarea value={origin.address} onChange={e => { originTyping.current = true; setOrigin({ address: e.target.value, lat: null, lng: null }); }} onBlur={() => setTimeout(() => setOriginSuggestions([]), 200)} placeholder="Digite o endereço de partida" rows={2} style={s.textarea} />
              {originSuggestions.length > 0 && (
                <div style={{ position:'absolute', top:'100%', left:0, right:0, zIndex:20, backgroundColor:C.card, borderRadius:12, marginTop:4, boxShadow:'0 8px 24px rgba(10,9,8,0.15)', border:`1.5px solid ${C.border}`, overflow:'hidden', maxHeight:200, overflowY:'auto' }}>
                  {originSuggestions.map((sg, i) => (
                    <button key={i} onMouseDown={e => e.preventDefault()} onClick={() => selectSuggestion('origin', sg)} style={{ width:'100%', padding:'10px 14px', background:'none', textAlign:'left', fontSize:12, fontFamily:fB, color:C.inkSoft, lineHeight:1.4, borderTop: i > 0 ? `1px solid ${C.border}` : 'none', display:'flex', alignItems:'flex-start', gap:8 }}>
                      <MapPin size={14} color={C.red} style={{ marginTop:2, flexShrink:0 }} strokeWidth={2.2} />
                      <span>{sg.display}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
            {origin.lat && <div style={{ marginTop:6 }}><span style={{ fontFamily:fM, fontSize:9, color:C.green, letterSpacing:'0.08em' }}>● coordenadas OK</span></div>}
            <button onClick={() => capture('origin')} disabled={loading.origin} className="km-press" style={{ ...s.btnRed, background: loading.origin ? C.borderDark : s.btnRed.background, boxShadow: loading.origin ? 'none' : s.btnRed.boxShadow, cursor: loading.origin ? 'not-allowed' : 'pointer' }}>
              {loading.origin ? <><Loader2 size={16} className="km-spin" /> Buscando…</> : <><MapPin size={16} strokeWidth={2.4} /> Capturar localização</>}
            </button>
          </div>

          <div style={{ display:'flex', alignItems:'center', gap:10, marginLeft:4, padding:'6px 0' }}>
            <ArrowDown size={12} color={C.inkFaded} strokeWidth={2.4} />
            <div style={{ flex:1, height:1, background:`repeating-linear-gradient(90deg, ${C.borderDark} 0 4px, transparent 4px 8px)` }} />
            <span style={{ fontFamily:fM, fontSize:9, color:C.inkFaded, letterSpacing:'0.16em', textTransform:'uppercase' }}>destino</span>
            <div style={{ flex:1, height:1, background:`repeating-linear-gradient(90deg, ${C.borderDark} 0 4px, transparent 4px 8px)` }} />
          </div>

          {/* Destination */}
          <div style={{ marginBottom:18 }}>
            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:10 }}>
              <label style={s.label}>
                <Flag size={14} color={C.black} fill={C.black} strokeWidth={2.4} />
                · Destino
              </label>
              {destination.address && <button onClick={() => clearLocation('destination')} style={{ fontFamily:fM, fontSize:10, color:C.inkFaded, background:'none', letterSpacing:'0.1em', textTransform:'uppercase' }}>clear</button>}
            </div>
            <div style={{ position:'relative' }}>
              <textarea value={destination.address} onChange={e => { destTyping.current = true; setDestination({ address: e.target.value, lat: null, lng: null }); }} onBlur={() => setTimeout(() => setDestSuggestions([]), 200)} placeholder="Digite o endereço de chegada" rows={2} style={s.textarea} />
              {destSuggestions.length > 0 && (
                <div style={{ position:'absolute', top:'100%', left:0, right:0, zIndex:20, backgroundColor:C.card, borderRadius:12, marginTop:4, boxShadow:'0 8px 24px rgba(10,9,8,0.15)', border:`1.5px solid ${C.border}`, overflow:'hidden', maxHeight:200, overflowY:'auto' }}>
                  {destSuggestions.map((sg, i) => (
                    <button key={i} onMouseDown={e => e.preventDefault()} onClick={() => selectSuggestion('destination', sg)} style={{ width:'100%', padding:'10px 14px', background:'none', textAlign:'left', fontSize:12, fontFamily:fB, color:C.inkSoft, lineHeight:1.4, borderTop: i > 0 ? `1px solid ${C.border}` : 'none', display:'flex', alignItems:'flex-start', gap:8 }}>
                      <Flag size={14} color={C.black} style={{ marginTop:2, flexShrink:0 }} strokeWidth={2.2} />
                      <span>{sg.display}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
            {destination.lat && <div style={{ marginTop:6 }}><span style={{ fontFamily:fM, fontSize:9, color:C.green, letterSpacing:'0.08em' }}>● coordenadas OK</span></div>}
            <button onClick={() => capture('destination')} disabled={loading.destination} className="km-press" style={{ ...s.btnBlack, background: loading.destination ? C.borderDark : C.black, boxShadow: loading.destination ? 'none' : s.btnBlack.boxShadow, cursor: loading.destination ? 'not-allowed' : 'pointer' }}>
              {loading.destination ? <><Loader2 size={16} className="km-spin" /> Buscando…</> : <><Flag size={16} strokeWidth={2.4} /> Capturar localização</>}
            </button>
          </div>

          {/* Odometer */}
          <div style={{ background:`linear-gradient(135deg, ${C.black} 0%, #1A1414 100%)`, borderRadius:18, padding:'18px 20px', marginBottom:16, display:'flex', alignItems:'center', justifyContent:'space-between', position:'relative', overflow:'hidden', border:`1px solid rgba(230,26,39,0.25)` }}>
            <div style={{ position:'absolute', top:-30, right:-30, width:120, height:120, borderRadius:60, background:`radial-gradient(circle, rgba(230,26,39,0.4), transparent 70%)`, pointerEvents:'none' }} />
            <div style={{ position:'absolute', inset:0, opacity:0.08, backgroundImage:'linear-gradient(rgba(255,255,255,0.4) 1px, transparent 1px)', backgroundSize:'100% 12px', pointerEvents:'none' }} />
            <div style={{ position:'relative' }}>
              <div style={{ fontFamily:fM, fontSize:9.5, letterSpacing:'0.26em', color:'rgba(255,255,255,0.55)', textTransform:'uppercase', fontWeight:500 }}>◊ Distância</div>
              {distance != null && (
                <div style={{ fontFamily:fM, fontSize:10, color:'#FF6B7A', marginTop:4, fontWeight:600, letterSpacing:'0.04em' }}>
                  {formatBRL(distance * RATE)}
                </div>
              )}
              {distanceLabel && (
                <div style={{ fontFamily:fM, fontSize:9.5, color:'rgba(255,255,255,0.4)', marginTop:3 }}>{distanceLabel}</div>
              )}
            </div>
            <div style={{ position:'relative' }}>
              {loading.distance
                ? <Loader2 size={24} color={C.red} className="km-spin" />
                : distance != null
                  ? <div style={{ display:'flex', alignItems:'baseline', gap:7 }}>
                      <span style={{ fontFamily:fM, fontSize:36, color:C.white, lineHeight:1, letterSpacing:'-0.02em', fontWeight:700, textShadow:`0 0 18px rgba(230,26,39,0.5)` }}>{distance.toFixed(2)}</span>
                      <span style={{ fontFamily:fD, fontSize:14, color:C.red, letterSpacing:'0.04em', fontWeight:700, fontStyle:'italic' }}>km</span>
                    </div>
                  : <div style={{ fontFamily:fM, fontSize:28, color:'rgba(255,255,255,0.22)', fontWeight:500 }}>--.--</div>
              }
            </div>
          </div>

          <button onClick={handleSave} disabled={!canSave} className="km-press" style={{ width:'100%', padding:'17px', borderRadius:16, background: canSave ? `linear-gradient(135deg, ${C.red} 0%, ${C.redDeep} 100%)` : C.borderDark, color:C.white, fontWeight:700, fontSize:16, fontFamily:fD, display:'flex', alignItems:'center', justifyContent:'center', gap:10, cursor: canSave ? 'pointer' : 'not-allowed', letterSpacing:'-0.01em', boxShadow: canSave ? '0 12px 28px rgba(230,26,39,0.36)' : 'none' }}>
            <Save size={16} strokeWidth={2.4} /> Registrar viagem
          </button>
        </section>

        {/* TAB SWITCHER */}
        <div style={{
          display:'flex', background:C.card, borderRadius:100, padding:5,
          marginBottom:14, border:`1px solid ${C.border}`,
          boxShadow:'0 2px 8px rgba(10,9,8,0.04)'
        }}>
          <button onClick={() => setActiveTab('history')} style={tabBtn(activeTab === 'history')}>
            <History size={14} strokeWidth={2.4} /> Histórico
          </button>
          <button onClick={() => setActiveTab('dashboard')} style={tabBtn(activeTab === 'dashboard')}>
            <TrendingUp size={14} strokeWidth={2.4} /> Dashboard
          </button>
        </div>

        {/* TAB CONTENT */}
        {activeTab === 'dashboard' ? (
          <Dashboard trips={trips} fD={fD} fM={fM} fB={fB} />
        ) : (
          <section style={{ ...s.card, marginBottom:0 }} className="km-fade">
            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:20 }}>
              <button onClick={() => setHistoryOpen(!historyOpen)} style={{ display:'flex', alignItems:'center', gap:12, background:'none', padding:0 }}>
                <div style={s.redBar} />
                <h2 style={{ fontFamily:fD, fontSize:26, color:C.ink, letterSpacing:'-0.03em', fontWeight:700, margin:0 }}>Histórico</h2>
                <span style={{ fontFamily:fM, fontSize:11, color:C.white, background:C.black, padding:'3px 9px', borderRadius:8, letterSpacing:'0.06em' }}>{String(trips.length).padStart(2,'0')}</span>
                {historyOpen ? <ChevronUp size={16} color={C.inkFaded} /> : <ChevronDown size={16} color={C.inkFaded} />}
              </button>
              <button onClick={exportToExcel} className="km-press" style={{ padding:'10px 14px', borderRadius:12, background:C.black, color:C.white, fontSize:12, fontWeight:600, fontFamily:fD, display:'flex', alignItems:'center', gap:7, letterSpacing:'0.01em' }}>
                <Download size={14} strokeWidth={2.4} /> Exportar
              </button>
            </div>

            {historyOpen && (
              trips.length === 0
                ? <p style={{ fontSize:13, color:C.inkFaded, textAlign:'center', padding:'24px 0', fontStyle:'italic', fontFamily:fD }}>Nenhuma viagem registrada ainda.</p>
                : <div>
                    {trips.map((trip, idx) => (
                      <article key={trip.id} style={{ borderTop: idx === 0 ? 'none' : `1px solid ${C.border}`, paddingTop: idx === 0 ? 0 : 16, paddingBottom:8, marginBottom:8 }}>
                        <div style={{ display:'flex', alignItems:'flex-start', justifyContent:'space-between', marginBottom:10 }}>
                          <span style={{ fontFamily:fM, fontSize:10.5, fontWeight:500, color:C.inkFaded, letterSpacing:'0.14em' }}>
                            #{String(trips.length - idx).padStart(3,'0')} · {trip.date}
                          </span>
                          <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                            {trip.km != null
                              ? <div style={{ display:'flex', flexDirection:'column', alignItems:'flex-end', gap:3 }}>
                                  <span style={{ fontFamily:fM, fontSize:11, fontWeight:700, background:`linear-gradient(135deg, ${C.red}, ${C.redDeep})`, color:C.white, padding:'4px 10px', borderRadius:100, letterSpacing:'0.04em', boxShadow:'0 4px 10px rgba(230,26,39,0.28)' }}>{trip.km.toFixed(2)} km</span>
                                  <span style={{ fontFamily:fM, fontSize:10, fontWeight:600, color:C.green, letterSpacing:'0.04em' }}>{formatBRL(trip.km * RATE)}</span>
                                </div>
                              : <span style={{ fontFamily:fM, fontSize:10, color:C.inkFaded, fontStyle:'italic' }}>sem medição</span>
                            }
                            <button onClick={() => deleteTrip(trip.id)} style={{ color:C.inkFaded, padding:3, background:'none' }}><Trash2 size={14} /></button>
                          </div>
                        </div>
                        <div style={{ fontSize:13, lineHeight:1.5, color:C.inkSoft, fontWeight:500 }}>
                          <div style={{ display:'flex', gap:10, alignItems:'flex-start' }}>
                            <span style={{ width:8, height:8, borderRadius:4, background:`radial-gradient(circle at 30% 30%, ${C.redBright}, ${C.redDeep})`, marginTop:6, flexShrink:0, display:'inline-block' }} />
                            <span>{trip.origin}</span>
                          </div>
                          <div style={{ marginLeft:3.5, height:12, width:1, background:C.borderDark, margin:'2px 0 2px 3.5px' }} />
                          <div style={{ display:'flex', gap:10, alignItems:'flex-start' }}>
                            <Flag size={12} color={C.black} fill={C.black} strokeWidth={2.4} style={{ marginTop:4, flexShrink:0 }} />
                            <span>{trip.destination}</span>
                          </div>
                        </div>

                        {/* MINI MAP */}
                        <MiniMap geometry={trip.geometry} fM={fM} />
                      </article>
                    ))}
                  </div>
            )}
          </section>
        )}

        <div style={{ marginTop:26, textAlign:'center' }}>
          <p style={{ fontFamily:fM, fontSize:10, color:C.inkFaded, lineHeight:1.7, letterSpacing:'0.12em', textTransform:'uppercase', margin:0 }}>
            ◊ Dados salvos no seu celular ◊<br />
            <span style={{ color:C.red }}>R$ 1,14 por km · exportar gera o excel</span>
          </p>
        </div>
      </div>
    </div>
  );
}
