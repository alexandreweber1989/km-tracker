import { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { MapPin, Flag, Save, Download, Trash2, Loader2, AlertCircle, X, Navigation, ChevronDown, ChevronUp, ArrowDown, DollarSign, Calendar, TrendingUp, History, Sun, Moon, Volume2, VolumeX, Vibrate } from 'lucide-react';
import * as XLSX from 'xlsx';

const RATE = 1.14;
const APP_VERSION = 'v3·1';

const SEED_TRIPS = [
  { id:'s1', date:'20/04/2026', origin:'Rua Attilio Ceccarelli, 90 - Jardim Rio Pequeno, São Paulo - SP, 05388-040', destination:'Rua dos Marianos, 349 - Centro, Osasco - SP, 06016-050', km:null, geometry:null },
  { id:'s2', date:'20/04/2026', origin:'Rua dos Marianos, 349 - Centro, Osasco - SP, 06016-050', destination:'Avenida Marechal Rondon, 199 - Centro, Osasco - SP, 06093-020', km:null, geometry:null },
  { id:'s3', date:'20/04/2026', origin:'Avenida Marechal Rondon, 165 - Centro, Osasco - SP, 06093-020', destination:'Rua Antônio Agú, 833 - Centro, Osasco - SP, 06013-000', km:null, geometry:null },
  { id:'s4', date:'20/04/2026', origin:'Rua Antônio Agú, 833 - Centro, Osasco - SP, 06013-000', destination:'Rua Minas Bogasian, 284 - Centro, Osasco - SP, 06013-010', km:null, geometry:null },
  { id:'s5', date:'20/04/2026', origin:'Rua Minas Bogasian, 284 - Centro, Osasco - SP, 06013-010', destination:'Rua Dona Primitiva Vianco, 589 - Centro, Osasco - SP, 06010-004', km:null, geometry:null },
  { id:'s6', date:'20/04/2026', origin:'Rua Dona Primitiva Vianco, 589 - Centro, Osasco - SP, 06010-004', destination:'Avenida João Batista, 11 - Centro, Osasco - SP, 06097-100', km:null, geometry:null },
  { id:'s7', date:'20/04/2026', origin:'Avenida João Batista, 11 - Centro, Osasco - SP, 06097-100', destination:'Avenida João Batista, 253 - Centro, Osasco - SP, 06090-100', km:null, geometry:null },
  { id:'s8', date:'20/04/2026', origin:'Avenida João Batista, 253 - Centro, Osasco - SP, 06090-100', destination:'Rua Fiorino Beltrano, 195 - Centro, Osasco - SP, 06097-040', km:null, geometry:null },
  { id:'s9', date:'20/04/2026', origin:'Rua Fiorino Beltrano, 195 - Centro, Osasco - SP, 06097-040', destination:'Rua Attilio Ceccarelli, 90 - Jardim Rio Pequeno, São Paulo - SP, 05388-040', km:null, geometry:null },
];

const STATE_MAP = {'Acre':'AC','Alagoas':'AL','Amapá':'AP','Amazonas':'AM','Bahia':'BA','Ceará':'CE','Distrito Federal':'DF','Espírito Santo':'ES','Goiás':'GO','Maranhão':'MA','Mato Grosso':'MT','Mato Grosso do Sul':'MS','Minas Gerais':'MG','Pará':'PA','Paraíba':'PB','Paraná':'PR','Pernambuco':'PE','Piauí':'PI','Rio de Janeiro':'RJ','Rio Grande do Norte':'RN','Rio Grande do Sul':'RS','Rondônia':'RO','Roraima':'RR','Santa Catarina':'SC','São Paulo':'SP','Sergipe':'SE','Tocantins':'TO'};
const MONTH_NAMES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const MONTH_FULL = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

// ═══ HELPERS ═══
function distMeters(lat1,lon1,lat2,lon2){const R=6371000,dLat=(lat2-lat1)*Math.PI/180,dLon=(lon2-lon1)*Math.PI/180,a=Math.sin(dLat/2)**2+Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;return 2*R*Math.asin(Math.sqrt(a));}
function haversineKm(a,b,c,d){return distMeters(a,b,c,d)/1000;}
function formatAddr(data,overrideNumber,isApprox){const a=data.address||{};const street=a.road||a.pedestrian||a.path||'';const number=overrideNumber||a.house_number||'';const neighborhood=a.suburb||a.neighbourhood||a.quarter||a.city_district||'';const city=a.city||a.town||a.village||a.municipality||'';const stateAbbr=STATE_MAP[a.state||'']||a.state||'';const postcode=a.postcode||'';let parts=[];if(street){let nl=number?(isApprox?`~${number}`:`${number}`):'';let s=nl?`${street}, ${nl}`:street;if(neighborhood)s+=` - ${neighborhood}`;parts.push(s);}else if(neighborhood)parts.push(neighborhood);let loc=city;if(stateAbbr)loc+=(loc?' - ':'')+stateAbbr;if(postcode)loc+=(loc?', ':'')+postcode;if(loc)parts.push(loc);return parts.length>0?parts.join(', '):(data.display_name||'');}
async function reverseLookup(lat,lng,zoom=19){const url=`https://nominatim.openstreetmap.org/reverse?lat=${lat}&lon=${lng}&format=json&accept-language=pt-BR&addressdetails=1&zoom=${zoom}`;const r=await fetch(url);if(!r.ok)throw new Error('Erro no serviço de endereços');return await r.json();}
async function searchNearbyForNumber(lat,lng,road){if(!road)return null;const delta=0.0015,viewbox=`${lng-delta},${lat-delta},${lng+delta},${lat+delta}`,q=encodeURIComponent(road);try{const r=await fetch(`https://nominatim.openstreetmap.org/search?q=${q}&format=json&addressdetails=1&accept-language=pt-BR&limit=20&bounded=1&viewbox=${viewbox}`);if(!r.ok)return null;const list=await r.json();if(!Array.isArray(list)||list.length===0)return null;const c=list.filter(i=>i.address?.house_number&&(i.address?.road===road||i.address?.pedestrian===road)).map(i=>({number:i.address.house_number,d:distMeters(lat,lng,parseFloat(i.lat),parseFloat(i.lon))})).sort((a,b)=>a.d-b.d);return c[0]?.number||null;}catch{return null;}}
async function getAddressWithNumber(lat,lng){const data=await reverseLookup(lat,lng,19);const direct=data?.address?.house_number;if(direct)return{data,number:direct,isApprox:false};const road=data?.address?.road||data?.address?.pedestrian;const nearby=await searchNearbyForNumber(lat,lng,road);if(nearby)return{data,number:nearby,isApprox:true};return{data,number:null,isApprox:false};}

const brlFmt=new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL',minimumFractionDigits:2,maximumFractionDigits:2});
const numFmt=new Intl.NumberFormat('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});

const TRIPS_KEY='km_trips_v1',THEME_KEY='coca_theme',SOUND_KEY='coca_sound',HAPTIC_KEY='coca_haptic';
const loadTrips=()=>{try{const r=localStorage.getItem(TRIPS_KEY);return r?JSON.parse(r):null;}catch{return null;}};
const saveTrips=(t)=>{try{localStorage.setItem(TRIPS_KEY,JSON.stringify(t));}catch{}};

// ═══ FEEDBACK ═══
let audioCtx;
function getAudioCtx(){if(typeof window==='undefined')return null;if(!audioCtx)try{audioCtx=new(window.AudioContext||window.webkitAudioContext)();}catch{return null;}return audioCtx;}
function playTone(freq,dur=0.04,type='sine',vol=0.15){const ctx=getAudioCtx();if(!ctx)return;const o=ctx.createOscillator(),g=ctx.createGain();o.type=type;o.frequency.value=freq;g.gain.setValueAtTime(0,ctx.currentTime);g.gain.linearRampToValueAtTime(vol,ctx.currentTime+0.005);g.gain.exponentialRampToValueAtTime(0.001,ctx.currentTime+dur);o.connect(g);g.connect(ctx.destination);o.start();o.stop(ctx.currentTime+dur);}
const SOUNDS={tap:()=>playTone(1000,0.025,'sine',0.10),success:()=>{playTone(660,0.06,'sine',0.12);setTimeout(()=>playTone(880,0.10,'sine',0.14),60);},error:()=>{playTone(400,0.08,'sawtooth',0.10);setTimeout(()=>playTone(280,0.12,'sawtooth',0.10),80);},swoosh:()=>playTone(800,0.05,'triangle',0.08)};
const HAPTICS={tap:10,success:[15,50,15],error:[50,30,50,30,50],swoosh:8};
function useFeedback(){return useCallback((key)=>{if(typeof window==='undefined')return;if(matchMedia('(prefers-reduced-motion: reduce)').matches)return;if(localStorage.getItem(SOUND_KEY)!=='off')try{SOUNDS[key]?.();}catch{};if(localStorage.getItem(HAPTIC_KEY)!=='off'&&'vibrate'in navigator)try{navigator.vibrate(HAPTICS[key]||10);}catch{};},[]);}

// ═══ THEME ═══
function getInitialTheme(){if(typeof window==='undefined')return'light';const s=localStorage.getItem(THEME_KEY);if(s==='dark'||s==='light')return s;return matchMedia('(prefers-color-scheme: dark)').matches?'dark':'light';}
function setTheme(next){document.documentElement.setAttribute('data-theme',next);document.documentElement.style.colorScheme=next;localStorage.setItem(THEME_KEY,next);}
function toggleTheme(current){const next=current==='dark'?'light':'dark';if(typeof document!=='undefined'&&document.startViewTransition)document.startViewTransition(()=>setTheme(next));else setTheme(next);return next;}

// ═══ MINI MAP ═══
function MiniMap({geometry}){
  if(!geometry||geometry.length<2)return(<div className="km-mini-empty"><span>◊ rota indisponível ◊</span></div>);
  let minLng=Infinity,maxLng=-Infinity,minLat=Infinity,maxLat=-Infinity;
  for(const[lng,lat]of geometry){if(lng<minLng)minLng=lng;if(lng>maxLng)maxLng=lng;if(lat<minLat)minLat=lat;if(lat>maxLat)maxLat=lat;}
  const padX=((maxLng-minLng)||0.001)*0.15,padY=((maxLat-minLat)||0.001)*0.20;
  minLng-=padX;maxLng+=padX;minLat-=padY;maxLat+=padY;
  const W=320,H=100,dataAspect=(maxLng-minLng)/(maxLat-minLat),viewAspect=W/H;
  if(dataAspect>viewAspect){const r=(maxLng-minLng)/viewAspect,m=(minLat+maxLat)/2;minLat=m-r/2;maxLat=m+r/2;}
  else{const r=(maxLat-minLat)*viewAspect,m=(minLng+maxLng)/2;minLng=m-r/2;maxLng=m+r/2;}
  const project=([lng,lat])=>[((lng-minLng)/(maxLng-minLng))*W,H-((lat-minLat)/(maxLat-minLat))*H];
  const points=geometry.map(project);
  const pathD=points.map(([x,y],i)=>`${i===0?'M':'L'}${x.toFixed(1)},${y.toFixed(1)}`).join(' ');
  const[sx,sy]=points[0],[ex,ey]=points[points.length-1];
  return(
    <div className="km-minimap">
      <svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="xMidYMid slice">
        <defs><pattern id="mapgrid" width="16" height="16" patternUnits="userSpaceOnUse"><path d="M 16 0 L 0 0 0 16" fill="none" stroke="var(--map-grid)" strokeWidth="0.5"/></pattern>
        <linearGradient id="routegrad" x1="0%" y1="0%" x2="100%" y2="0%"><stop offset="0%" stopColor="var(--coca-red)"/><stop offset="100%" stopColor="var(--text-primary)"/></linearGradient></defs>
        <rect width={W} height={H} fill="url(#mapgrid)"/>
        <path d={pathD} stroke="var(--coca-red-glow)" strokeWidth="6" fill="none" strokeLinecap="round" strokeLinejoin="round"/>
        <path d={pathD} stroke="url(#routegrad)" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round" className="km-route-dash"/>
        <circle cx={sx} cy={sy} r="5" fill="var(--bg-base)" stroke="var(--coca-red)" strokeWidth="2"/>
        <circle cx={sx} cy={sy} r="2" fill="var(--coca-red)"/>
        <rect x={ex-4} y={ey-4} width="8" height="8" fill="var(--text-primary)" stroke="var(--bg-base)" strokeWidth="1.5"/>
      </svg>
      <div className="km-minimap-label"><span>◊ ROTA</span><span>{geometry.length} pts</span></div>
    </div>
  );
}

// ═══ ACTIVITY RINGS ═══
function ActivityRings({kmProgress,moneyProgress,daysProgress}){
  const rings=[{progress:kmProgress,color:'var(--coca-red)',radius:70,label:'KM'},{progress:moneyProgress,color:'var(--text-primary)',radius:54,label:'R$'},{progress:daysProgress,color:'var(--text-secondary)',radius:38,label:'DIAS'}];
  const stroke=11;
  return(
    <div className="km-rings">
      <svg viewBox="0 0 200 200" className="km-rings-svg">
        {rings.map((r,i)=>{const c=2*Math.PI*r.radius;return(<g key={i}><circle r={r.radius} cx="100" cy="100" fill="none" stroke="var(--border-subtle)" strokeWidth={stroke}/><circle r={r.radius} cx="100" cy="100" fill="none" stroke={r.color} strokeWidth={stroke} strokeDasharray={c} strokeDashoffset={c*(1-Math.min(r.progress,1))} strokeLinecap="round" transform="rotate(-90 100 100)" style={{transition:'stroke-dashoffset 1.2s cubic-bezier(.32,.72,0,1)'}}/></g>);})}
      </svg>
      <div className="km-rings-legend">
        {rings.map((r,i)=>(<div key={i} className="km-rings-legend-row"><span className="km-rings-dot" style={{background:r.color}}/><span className="km-rings-legend-label">{r.label}</span><span className="km-rings-legend-pct">{Math.round(r.progress*100)}%</span></div>))}
      </div>
    </div>
  );
}

// ═══ SPARKLINE ═══
function Sparkline({values,height=36}){
  if(!values||values.length===0)return null;
  const W=280,H=height,max=Math.max(...values,0.01),step=W/Math.max(values.length-1,1);
  const points=values.map((v,i)=>[i*step,H-(v/max)*(H-4)-2]);
  const pathD=points.map(([x,y],i)=>`${i===0?'M':'L'}${x.toFixed(1)},${y.toFixed(1)}`).join(' ');
  const areaD=pathD+` L${W},${H} L0,${H} Z`;
  return(
    <svg viewBox={`0 0 ${W} ${H}`} className="km-sparkline" preserveAspectRatio="none">
      <defs><linearGradient id="sparkfill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="var(--coca-red)" stopOpacity="0.3"/><stop offset="100%" stopColor="var(--coca-red)" stopOpacity="0"/></linearGradient></defs>
      <path d={areaD} fill="url(#sparkfill)"/>
      <path d={pathD} fill="none" stroke="var(--coca-red)" strokeWidth="1.5" strokeLinejoin="round" strokeLinecap="round"/>
    </svg>
  );
}

// ═══ BIG CURRENCY ═══
function BigCurrency({value,size='xl'}){
  const parts=brlFmt.formatToParts(value);
  let symbol='',integer='',decimal='';
  for(const p of parts){if(p.type==='currency')symbol=p.value;else if(p.type==='integer'||p.type==='group')integer+=p.value;else if(p.type==='decimal')decimal+=p.value;else if(p.type==='fraction')decimal+=p.value;}
  return(<span className={`km-big-currency km-big-currency--${size}`}><span className="km-big-currency-symbol">{symbol}</span><span className="km-big-currency-int">{integer}</span><span className="km-big-currency-dec">{decimal}</span></span>);
}

// ═══ DASHBOARD ═══
function Dashboard({trips}){
  const stats=useMemo(()=>{
    const totalKm=trips.reduce((s,t)=>s+(t.km||0),0);
    const measured=trips.filter(t=>t.km!=null);
    const byDay={};for(const t of measured){if(!byDay[t.date])byDay[t.date]={km:0,count:0};byDay[t.date].km+=t.km;byDay[t.date].count+=1;}
    const days=Object.entries(byDay).map(([date,d])=>({date,...d})).sort((a,b)=>{const p=s=>{const[d,m,y]=s.split('/');return new Date(y,m-1,d);};return p(b.date)-p(a.date);});
    const byMonth={};for(const t of measured){const[d,m,y]=t.date.split('/');const key=`${m}/${y}`;if(!byMonth[key])byMonth[key]={km:0,count:0,days:new Set()};byMonth[key].km+=t.km;byMonth[key].count+=1;byMonth[key].days.add(t.date);}
    const months=Object.entries(byMonth).map(([key,d])=>({key,...d,daysCount:d.days.size})).sort((a,b)=>{const[ma,ya]=a.key.split('/');const[mb,yb]=b.key.split('/');return new Date(yb,mb-1)-new Date(ya,ma-1);});
    const maxDayKm=Math.max(...days.map(d=>d.km),0.01);
    const avgPerDay=days.length>0?totalKm/days.length:0;
    const bestDay=days.reduce((a,b)=>(a?.km||0)>b.km?a:b,null);
    const now=new Date();const currentKey=`${String(now.getMonth()+1).padStart(2,'0')}/${now.getFullYear()}`;
    const currentMonth=byMonth[currentKey];const daysInMonth=new Date(now.getFullYear(),now.getMonth()+1,0).getDate();const today=now.getDate();
    let projection=null;if(currentMonth&&today>0){const projKm=(currentMonth.km/Math.min(today,daysInMonth))*daysInMonth;projection={km:projKm,money:projKm*RATE};}
    const last14=[];for(let i=13;i>=0;i--){const d=new Date(now);d.setDate(now.getDate()-i);const key=`${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;last14.push(byDay[key]?.km||0);}
    return{totalKm,measuredCount:measured.length,days,months,maxDayKm,avgPerDay,bestDay,projection,last14,currentMonth};
  },[trips]);
  const totalEarning=stats.totalKm*RATE;
  const now=new Date();

  return(
    <div className="km-dash km-fade-up">
      {/* ROW 1: Hero + Projection (side by side on desktop) */}
      <div className="km-dash-top-grid">
        {/* HERO METRIC */}
        <section className="km-card km-card--hero">
          <div className="km-aurora"/>
          <header className="km-card-head"><span className="km-mono km-mono-label">[ TOTAL · ALL-TIME ]</span><span className="km-mono km-mono-label">{String(trips.length).padStart(3,'0')}/REC</span></header>
          <div className="km-hero-amount"><BigCurrency value={totalEarning} size="xl"/></div>
          <div className="km-hero-meta"><span className="km-mono">{numFmt.format(stats.totalKm)} KM</span><span className="km-divider-vert"/><span className="km-mono km-muted">× R$ 1,14/KM</span></div>
          {stats.last14.some(v=>v>0)&&(<div className="km-hero-spark"><Sparkline values={stats.last14} height={32}/><span className="km-mono km-mono-tiny">14d</span></div>)}
        </section>

        {/* PROJECTION */}
        {stats.projection&&(
          <section className="km-card km-card--projection">
            <div className="km-mono km-mono-label" style={{marginBottom:8}}>◊ PROJEÇÃO · {MONTH_FULL[now.getMonth()].toUpperCase()} {now.getFullYear()}</div>
            <div className="km-projection-content">
              <BigCurrency value={stats.projection.money} size="md"/>
              <div className="km-mono km-muted km-mono-tiny" style={{marginTop:6}}>{numFmt.format(stats.projection.km)} KM ESTIMADOS</div>
            </div>
          </section>
        )}
      </div>

      {/* ROW 2: Rings + Stats (side by side on tablet+) */}
      <div className="km-dash-mid-grid">
        <section className="km-card km-card--rings">
          <header className="km-card-head"><span className="km-mono km-mono-label">◊ ATIVIDADE · MÊS</span></header>
          <ActivityRings kmProgress={Math.min((stats.currentMonth?.km||0)/1500,1)} moneyProgress={Math.min((stats.currentMonth?.km||0)*RATE/2000,1)} daysProgress={Math.min((stats.currentMonth?.daysCount||0)/22,1)}/>
        </section>

        <div className="km-stats-grid">
          <div className="km-stat-tile"><span className="km-mono km-mono-label">[ MÉDIA / DIA ]</span><span className="km-stat-num">{numFmt.format(stats.avgPerDay)}</span><span className="km-mono km-muted km-mono-tiny">KM</span></div>
          <div className="km-stat-tile"><span className="km-mono km-mono-label">[ MELHOR DIA ]</span><span className="km-stat-num">{stats.bestDay?numFmt.format(stats.bestDay.km):'0'}</span><span className="km-mono km-muted km-mono-tiny">{stats.bestDay?.date||'—'}</span></div>
        </div>
      </div>

      {/* PULL QUOTE */}
      {stats.bestDay&&stats.totalKm>0&&(
        <blockquote className="km-pullquote"><p><span className="km-serif km-italic">Maior consumo da rota em </span><span className="km-mono km-pullquote-data">{stats.bestDay.date.slice(3)}</span><span className="km-serif km-italic">.</span></p><cite className="km-mono km-mono-tiny km-muted">— REGISTRO INTERNO · DRIVE LOG</cite></blockquote>
      )}

      {/* ROW 3: By Day + By Month (side by side on desktop) */}
      <div className="km-dash-bottom-grid">
        <section className="km-card">
          <header className="km-card-head km-card-head--bordered"><div className="km-row-tight"><Calendar size={14} strokeWidth={2.4}/><span className="km-mono km-mono-label">POR DIA</span></div><span className="km-mono km-mono-label">{String(stats.days.length).padStart(2,'0')}</span></header>
          {stats.days.length===0?(<p className="km-empty">Sem dados ainda.</p>):(
            <ul className="km-rows">{stats.days.map((d)=>{const earning=d.km*RATE;const pct=(d.km/stats.maxDayKm)*100;return(
              <li key={d.date} className="km-row"><div className="km-row-main"><div><div className="km-mono km-row-date">{d.date}</div><div className="km-mono km-mono-tiny km-muted">{d.count} {d.count===1?'VIAGEM':'VIAGENS'}</div></div><div className="km-row-right"><div className="km-mono km-row-km">{numFmt.format(d.km)} KM</div><div className="km-row-money">{brlFmt.format(earning)}</div></div></div><div className="km-bar"><div className="km-bar-fill" style={{width:`${pct}%`}}/></div></li>
            );})}</ul>
          )}
        </section>

        <section className="km-card">
          <header className="km-card-head km-card-head--bordered"><div className="km-row-tight"><TrendingUp size={14} strokeWidth={2.4}/><span className="km-mono km-mono-label">POR MÊS</span></div><span className="km-mono km-mono-label">{String(stats.months.length).padStart(2,'0')}</span></header>
          {stats.months.length===0?(<p className="km-empty">Sem dados ainda.</p>):(
            <ul className="km-month-rows">{stats.months.map((m)=>{const[mm,yy]=m.key.split('/');return(
              <li key={m.key} className="km-month-row"><div><div className="km-month-title">{MONTH_NAMES[parseInt(mm)-1]} <span className="km-mono km-muted">·{yy}</span></div><div className="km-mono km-mono-tiny km-muted">{m.count} VIAGENS · {numFmt.format(m.km)} KM · {m.daysCount} DIAS</div></div><div className="km-month-money">{brlFmt.format(m.km*RATE)}</div></li>
            );})}</ul>
          )}
        </section>
      </div>
    </div>
  );
}

// ═══ SPLASH ═══
function Splash({exiting}){return(
  <div className={`km-splash ${exiting?'km-splash--exit':''}`}>
    <div className="km-splash-grid"/><div className="km-splash-content">
    <div className="km-splash-mono">[ INIT · {APP_VERSION} ]</div>
    <h1 className="km-splash-title"><span>DRIVE</span><span className="km-splash-italic">LOG</span></h1>
    <div className="km-splash-bar"><div className="km-splash-bar-fill"/></div>
    <div className="km-splash-mono km-splash-mono--small">SYS · BOOT · OK</div>
  </div></div>
);}

// ═══ MAIN APP ═══
export default function KmTracker(){
  const[origin,setOrigin]=useState({address:'',lat:null,lng:null});
  const[destination,setDestination]=useState({address:'',lat:null,lng:null});
  const[distance,setDistance]=useState(null);
  const[distanceLabel,setDistanceLabel]=useState('');
  const[routeGeometry,setRouteGeometry]=useState(null);
  const[trips,setTrips]=useState([]);
  const[loading,setLoading]=useState({origin:false,destination:false,distance:false});
  const[error,setError]=useState(null);
  const[success,setSuccess]=useState(null);
  const[activeTab,setActiveTab]=useState('history');
  const[time,setTime]=useState('');
  const[theme,setThemeState]=useState('light');
  const[soundOn,setSoundOn]=useState(true);
  const[hapticOn,setHapticOn]=useState(true);
  const[splashExiting,setSplashExiting]=useState(false);
  const[splashGone,setSplashGone]=useState(false);
  const[topLoading,setTopLoading]=useState(false);
  const[originSuggestions,setOriginSuggestions]=useState([]);
  const[destSuggestions,setDestSuggestions]=useState([]);
  const[showOriginDrop,setShowOriginDrop]=useState(false);
  const[showDestDrop,setShowDestDrop]=useState(false);
  const originDebounce=useRef(null);
  const destDebounce=useRef(null);
  const initialized=useRef(false);
  const feedback=useFeedback();

  useEffect(()=>{setThemeState(getInitialTheme());setSoundOn(localStorage.getItem(SOUND_KEY)!=='off');setHapticOn(localStorage.getItem(HAPTIC_KEY)!=='off');const t1=setTimeout(()=>setSplashExiting(true),1100);const t2=setTimeout(()=>setSplashGone(true),1700);return()=>{clearTimeout(t1);clearTimeout(t2);};},[]);
  useEffect(()=>{const update=()=>{const d=new Date();setTime(`${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`);};update();const id=setInterval(update,1000);return()=>clearInterval(id);},[]);
  useEffect(()=>{const loaded=loadTrips();if(loaded&&Array.isArray(loaded))setTrips(loaded);else{setTrips(SEED_TRIPS);saveTrips(SEED_TRIPS);}},[]);
  useEffect(()=>{if(trips.length>0)saveTrips(trips);},[trips]);
  useEffect(()=>{if(initialized.current)return;initialized.current=true;const fonts=document.createElement('link');fonts.href='https://fonts.googleapis.com/css2?family=Geist:wght@300;400;500;600;700;800;900&family=Geist+Mono:wght@400;500;700&family=Instrument+Serif:ital@0;1&display=swap';fonts.rel='stylesheet';document.head.appendChild(fonts);const style=document.createElement('style');style.textContent=CSS;document.head.appendChild(style);},[]);

  useEffect(()=>{if(origin.lat==null||destination.lat==null){setDistance(null);setDistanceLabel('');setRouteGeometry(null);return;}let cancelled=false;(async()=>{setLoading(s=>({...s,distance:true}));try{const r=await fetch(`https://router.project-osrm.org/route/v1/driving/${origin.lng},${origin.lat};${destination.lng},${destination.lat}?overview=simplified&geometries=geojson`);const data=await r.json();if(cancelled)return;if(data.routes?.length>0){const route=data.routes[0];setDistance(route.distance/1000);setDistanceLabel('rota de carro');setRouteGeometry(route.geometry?.coordinates||null);}else throw new Error();}catch{if(cancelled)return;setDistance(haversineKm(origin.lat,origin.lng,destination.lat,destination.lng));setDistanceLabel('linha reta · rota indisponível');setRouteGeometry([[origin.lng,origin.lat],[destination.lng,destination.lat]]);}finally{if(!cancelled)setLoading(s=>({...s,distance:false}));}})();return()=>{cancelled=true;};},[origin.lat,origin.lng,destination.lat,destination.lng]);

  async function searchAddress(query,which){if(query.length<3){if(which==='origin'){setOriginSuggestions([]);setShowOriginDrop(false);}else{setDestSuggestions([]);setShowDestDrop(false);}return;}const typedNum=query.match(/[\s,]+(\d{1,5})(?:\s*[-,]|\s*$)/);const num=typedNum?typedNum[1]:null;try{const url=`https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(query)}&format=json&addressdetails=1&accept-language=pt-BR&countrycodes=br&limit=5`;const r=await fetch(url);if(!r.ok)return;const list=await r.json();const results=list.map(item=>{const a=item.address||{};const street=a.road||a.pedestrian||a.path||'';const houseNum=a.house_number||num||'';const neighborhood=a.suburb||a.neighbourhood||a.quarter||a.city_district||'';const city=a.city||a.town||a.village||a.municipality||'';const stateAbbr=STATE_MAP[a.state||'']||a.state||'';let label=street;if(houseNum)label+=`, ${houseNum}`;if(neighborhood)label+=` - ${neighborhood}`;if(city)label+=`, ${city}`;if(stateAbbr)label+=` - ${stateAbbr}`;return{label:label||item.display_name,lat:parseFloat(item.lat),lng:parseFloat(item.lon)};});if(which==='origin'){setOriginSuggestions(results);setShowOriginDrop(results.length>0);}else{setDestSuggestions(results);setShowDestDrop(results.length>0);}}catch{}}
  function handleAddressInput(e,which){const val=e.target.value;const setter=which==='origin'?setOrigin:setDestination;setter(prev=>({...prev,address:val,lat:null,lng:null}));const debRef=which==='origin'?originDebounce:destDebounce;clearTimeout(debRef.current);debRef.current=setTimeout(()=>searchAddress(val,which),400);}
  function selectSuggestion(item,which){const setter=which==='origin'?setOrigin:setDestination;setter({address:item.label,lat:item.lat,lng:item.lng});if(which==='origin'){setOriginSuggestions([]);setShowOriginDrop(false);}else{setDestSuggestions([]);setShowDestDrop(false);}feedback('tap');}
  async function capture(which){setError(null);setLoading(s=>({...s,[which]:true}));setTopLoading(true);feedback('tap');try{if(!navigator.geolocation)throw new Error('Geolocalização não disponível');const pos=await new Promise((res,rej)=>navigator.geolocation.getCurrentPosition(p=>res({lat:p.coords.latitude,lng:p.coords.longitude}),e=>rej(new Error(e.code===1?'Permissão negada. Habilite o GPS.':e.code===2?'GPS indisponível.':e.code===3?'Timeout. Tente novamente.':'Erro de localização')),{enableHighAccuracy:true,timeout:20000,maximumAge:0}));const result=await getAddressWithNumber(pos.lat,pos.lng);const address=formatAddr(result.data,result.number,result.isApprox);(which==='origin'?setOrigin:setDestination)({address,lat:pos.lat,lng:pos.lng});feedback('success');if(result.number&&result.isApprox){setSuccess('Número aproximado · pode ajustar');setTimeout(()=>setSuccess(null),2600);}else if(!result.number){setSuccess('Sem número · adicione manual');setTimeout(()=>setSuccess(null),2600);}}catch(e){setError(e.message);feedback('error');}finally{setLoading(s=>({...s,[which]:false}));setTopLoading(false);}}
  function clearLocation(w){feedback('tap');(w==='origin'?setOrigin:setDestination)({address:'',lat:null,lng:null});if(w==='origin'){setOriginSuggestions([]);setShowOriginDrop(false);}else{setDestSuggestions([]);setShowDestDrop(false);}}
  function handleSave(){if(!origin.address.trim()||!destination.address.trim()){setError('Preencha origem e destino');feedback('error');return;}const d=new Date();const newTrip={id:'t'+Date.now(),date:`${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`,origin:origin.address.trim(),destination:destination.address.trim(),km:distance!=null?Number(distance.toFixed(2)):null,kmLabel:distanceLabel,geometry:routeGeometry};setTrips(prev=>[newTrip,...prev]);setOrigin({address:'',lat:null,lng:null});setDestination({address:'',lat:null,lng:null});setDistance(null);setDistanceLabel('');setRouteGeometry(null);setError(null);setSuccess('Viagem registrada');setTimeout(()=>setSuccess(null),2200);feedback('success');}
  function deleteTrip(id){if(window.confirm('Remover esta viagem?')){setTrips(prev=>prev.filter(t=>t.id!==id));feedback('swoosh');}}
  function exportToExcel(){if(!trips.length){setError('Nenhuma viagem para exportar');feedback('error');return;}const sorted=[...trips].sort((a,b)=>{const parse=s=>{const[d,m,y]=s.split('/');return new Date(y,m-1,d);};return parse(a.date)-parse(b.date);});const data=[['Data','Origem','Destino'],...sorted.map(t=>[t.date,t.origin,t.destination])];const ws=XLSX.utils.aoa_to_sheet(data);ws['!cols']=[{wch:12},{wch:60},{wch:60}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'kmadicional');XLSX.writeFile(wb,'lançamento_de_Km.xlsx');setSuccess('Planilha exportada');setTimeout(()=>setSuccess(null),2200);feedback('success');}
  function handleToggleTheme(){feedback('tap');setThemeState(toggleTheme(theme));}
  function handleToggleSound(){const next=!soundOn;setSoundOn(next);localStorage.setItem(SOUND_KEY,next?'on':'off');if(next)feedback('tap');}
  function handleToggleHaptic(){const next=!hapticOn;setHapticOn(next);localStorage.setItem(HAPTIC_KEY,next?'on':'off');if(next&&'vibrate'in navigator)navigator.vibrate(15);}
  function handleTabSwitch(t){if(t===activeTab)return;feedback('swoosh');if(document.startViewTransition)document.startViewTransition(()=>setActiveTab(t));else setActiveTab(t);}

  const totalKm=trips.reduce((s,t)=>s+(t.km||0),0);
  const totalEarning=totalKm*RATE;
  const tripsWithKm=trips.filter(t=>t.km!=null).length;
  const canSave=origin.address.trim()&&destination.address.trim();
  const marqueeText=' · DRIVE LOG · COCA-COLA BR · FROTA · '+numFmt.format(totalKm)+' KM · '+brlFmt.format(totalEarning)+' · '+String(trips.length).padStart(3,'0')+' VIAGENS';

  return(
    <div className="km-app" data-theme={theme}>
      {!splashGone&&<Splash exiting={splashExiting}/>}
      {topLoading&&<div className="km-toploader"/>}

      {/* STATUS BAR */}
      <div className="km-statusbar"><div className="km-statusbar-inner">
        <span className="km-mono km-mono-tiny"><span className="km-status-dot"/> SYS · ONLINE</span>
        <span className="km-mono km-mono-tiny km-muted">{APP_VERSION} · DRIVE LOG</span>
        <span className="km-mono km-mono-tiny">{time}</span>
      </div></div>

      {/* HEADER */}
      <header className="km-header">
        <div className="km-aurora-bg"/><div className="km-grid-overlay"/>
        <div className="km-header-inner">
          <div className="km-header-top">
            <div>
              <div className="km-mono km-mono-tiny km-muted">[ {APP_VERSION} ] · MOTORISTA</div>
              <h1 className="km-h1"><span>CONTROLE</span><span className="km-h1-italic">de KM</span></h1>
            </div>
            <div className="km-header-actions">
              <button onClick={handleToggleTheme} className="km-icon-btn" aria-label="Tema">{theme==='dark'?<Sun size={16} strokeWidth={2.4}/>:<Moon size={16} strokeWidth={2.4}/>}</button>
              <button onClick={handleToggleSound} className="km-icon-btn" aria-label="Som">{soundOn?<Volume2 size={16} strokeWidth={2.4}/>:<VolumeX size={16} strokeWidth={2.4}/>}</button>
              <button onClick={handleToggleHaptic} className="km-icon-btn" aria-label="Vibração"><Vibrate size={16} strokeWidth={2.4} style={{opacity:hapticOn?1:0.35}}/></button>
            </div>
          </div>
          {/* HERO */}
          <div className="km-hero">
            <div className="km-mono km-mono-label"><span className="km-line"/>VALOR A RECEBER</div>
            <div className="km-hero-value"><BigCurrency value={totalEarning} size="hero"/></div>
            <div className="km-hero-meta-row">
              <span className="km-pill"><span className="km-status-dot km-status-dot--green"/><span className="km-mono km-mono-tiny">{numFmt.format(totalKm)} KM TOTAIS</span></span>
              <span className="km-mono km-mono-tiny km-muted">{String(trips.length).padStart(3,'0')} TRIPS · {String(tripsWithKm).padStart(3,'0')} MED</span>
            </div>
          </div>
        </div>
        <svg viewBox="0 0 400 60" preserveAspectRatio="none" className="km-wave"><path d="M0,40 C60,10 120,55 200,30 C280,5 340,50 400,25 L400,60 L0,60 Z" fill="var(--bg-base)"/><path d="M0,40 C60,10 120,55 200,30 C280,5 340,50 400,25" fill="none" stroke="rgba(255,255,255,0.4)" strokeWidth="1"/></svg>
      </header>

      {/* CONTENT */}
      <main className="km-main">
        <div className="km-content-grid">
          {/* LEFT SIDE: New Trip (on desktop becomes a sticky sidebar) */}
          <div className="km-content-left">
            {error&&(<div className="km-alert km-alert--err km-fade-in"><AlertCircle size={18} strokeWidth={2.4}/><p>{error}</p><button onClick={()=>setError(null)}><X size={14}/></button></div>)}
            {success&&(<div className="km-alert km-alert--ok km-fade-in"><span className="km-status-dot km-status-dot--green"/><p>{success}</p></div>)}

            <section className="km-card">
              <header className="km-card-head km-card-head--bordered"><div className="km-row-tight"><span className="km-redbar"/><h2 className="km-h2">Nova viagem</h2></div><span className="km-mono km-mono-tiny km-muted">001/REC</span></header>
              <div className="km-field"><div className="km-field-head"><label htmlFor="origin-address" className="km-mono km-mono-label"><span className="km-dot km-dot--red"/>· ORIGEM</label>{origin.address&&<button onClick={()=>clearLocation('origin')} className="km-textbtn" aria-label="Limpar endereço de origem">CLEAR</button>}</div><div className="km-autocomplete"><textarea id="origin-address" value={origin.address} onChange={e=>handleAddressInput(e,'origin')} onFocus={()=>originSuggestions.length>0&&setShowOriginDrop(true)} onBlur={()=>setTimeout(()=>setShowOriginDrop(false),200)} placeholder="Digite o endereço de partida" rows={2} className="km-textarea" aria-label="Endereço de origem" aria-describedby="origin-help" aria-autocomplete="list" aria-controls="origin-suggestions"/>{showOriginDrop&&originSuggestions.length>0&&(<ul id="origin-suggestions" className="km-suggestions" role="listbox">{originSuggestions.map((s,i)=>(<li key={i} onMouseDown={()=>selectSuggestion(s,'origin')} className="km-suggestion-item" role="option">{s.label}</li>))}</ul>)}</div><div id="origin-help" className="km-mono km-mono-tiny km-muted" style={{marginTop:'4px'}}>Comece a digitar para autocomplete ou clique abaixo para usar GPS</div><button onClick={()=>capture('origin')} disabled={loading.origin} className="km-btn km-btn--red km-btn--block km-press" aria-label={loading.origin?'Buscando localização de origem':'Capturar localização de origem'}>{loading.origin?<><Loader2 size={15} className="km-spin"/> BUSCANDO…</>:<><MapPin size={15} strokeWidth={2.4}/> CAPTURAR LOCALIZAÇÃO</>}</button></div>
              <div className="km-connector"><ArrowDown size={11} strokeWidth={2.4}/><div className="km-dashed"/><span className="km-mono km-mono-tiny">DESTINO</span><div className="km-dashed"/></div>
              <div className="km-field"><div className="km-field-head"><label htmlFor="dest-address" className="km-mono km-mono-label"><Flag size={11} strokeWidth={2.4} fill="currentColor"/>· DESTINO</label>{destination.address&&<button onClick={()=>clearLocation('destination')} className="km-textbtn" aria-label="Limpar endereço de destino">CLEAR</button>}</div><div className="km-autocomplete"><textarea id="dest-address" value={destination.address} onChange={e=>handleAddressInput(e,'destination')} onFocus={()=>destSuggestions.length>0&&setShowDestDrop(true)} onBlur={()=>setTimeout(()=>setShowDestDrop(false),200)} placeholder="Digite o endereço de chegada" rows={2} className="km-textarea" aria-label="Endereço de destino" aria-describedby="dest-help" aria-autocomplete="list" aria-controls="dest-suggestions"/>{showDestDrop&&destSuggestions.length>0&&(<ul id="dest-suggestions" className="km-suggestions" role="listbox">{destSuggestions.map((s,i)=>(<li key={i} onMouseDown={()=>selectSuggestion(s,'destination')} className="km-suggestion-item" role="option">{s.label}</li>))}</ul>)}</div><div id="dest-help" className="km-mono km-mono-tiny km-muted" style={{marginTop:'4px'}}>Comece a digitar para autocomplete ou clique abaixo para usar GPS</div><button onClick={()=>capture('destination')} disabled={loading.destination} className="km-btn km-btn--ink km-btn--block km-press" aria-label={loading.destination?'Buscando localização de destino':'Capturar localização de destino'}>{loading.destination?<><Loader2 size={15} className="km-spin"/> BUSCANDO…</>:<><Flag size={15} strokeWidth={2.4}/> CAPTURAR LOCALIZAÇÃO</>}</button></div>
              <div className="km-odometer" role="region" aria-live="polite" aria-label="Informações de distância"><div className="km-odometer-glow"/><div className="km-odometer-grid"/><div className="km-odometer-info"><div className="km-mono km-mono-label km-odometer-label">◊ DISTÂNCIA</div>{distance!=null&&<div className="km-odometer-money" aria-label={`Valor a receber: ${brlFmt.format(distance*RATE)}`}>{brlFmt.format(distance*RATE)}</div>}{distanceLabel&&<div className="km-mono km-mono-tiny km-odometer-sublabel">{distanceLabel}</div>}</div><div className="km-odometer-display">{loading.distance?<Loader2 size={28} className="km-spin" style={{color:'var(--coca-red)'}} aria-label="Calculando distância"/>:distance!=null?<div className="km-odometer-num"><span aria-label={`${numFmt.format(distance)} quilômetros`}>{numFmt.format(distance)}</span><span className="km-odometer-unit">KM</span></div>:<div className="km-odometer-empty">--.--</div>}</div></div>
              <button onClick={handleSave} disabled={!canSave} className="km-btn km-btn--save km-btn--block km-press" aria-label={!canSave?'Preencha origem e destino para registrar viagem':'Registrar viagem'}><Save size={16} strokeWidth={2.4}/> REGISTRAR VIAGEM</button>
            </section>
          </div>

          {/* RIGHT SIDE: Tabs + Content */}
          <div className="km-content-right">
            <div className="km-tabs">
              <button onClick={()=>handleTabSwitch('history')} className={`km-tab ${activeTab==='history'?'km-tab--active':''}`}><History size={14} strokeWidth={2.4}/> HISTÓRICO</button>
              <button onClick={()=>handleTabSwitch('dashboard')} className={`km-tab ${activeTab==='dashboard'?'km-tab--active':''}`}><TrendingUp size={14} strokeWidth={2.4}/> DASHBOARD</button>
            </div>

            <div className="km-tabcontent" key={activeTab}>
              {activeTab==='dashboard'?(<Dashboard trips={trips}/>):(
                <section className="km-card km-fade-up">
                  <header className="km-card-head km-card-head--bordered"><div className="km-row-tight"><span className="km-redbar"/><h2 className="km-h2">Histórico</h2><span className="km-mono km-mono-tiny km-pill-dark">{String(trips.length).padStart(2,'0')}</span></div><button onClick={exportToExcel} className="km-btn km-btn--ink km-btn--small km-press"><Download size={13} strokeWidth={2.4}/> EXPORTAR</button></header>
                  {trips.length===0?(<p className="km-empty">Nenhuma viagem registrada ainda.</p>):(
                    <ul className="km-trip-list">{trips.map((trip,idx)=>(
                      <li key={trip.id} className="km-trip">
                        <div className="km-trip-head"><span className="km-mono km-mono-tiny">#{String(trips.length-idx).padStart(3,'0')} <span className="km-muted">·</span> {trip.date}</span><div className="km-trip-head-right">{trip.km!=null?(<div className="km-trip-km"><span className="km-trip-km-pill">{numFmt.format(trip.km)} KM</span><span className="km-mono km-mono-tiny km-trip-km-money">{brlFmt.format(trip.km*RATE)}</span></div>):(<span className="km-mono km-mono-tiny km-muted km-italic">SEM MEDIÇÃO</span>)}<button onClick={()=>deleteTrip(trip.id)} className="km-icon-btn-tiny"><Trash2 size={13}/></button></div></div>
                        <div className="km-trip-body"><div className="km-trip-line"><span className="km-dot km-dot--red"/><span>{trip.origin}</span></div><div className="km-trip-vline"/><div className="km-trip-line"><Flag size={11} fill="currentColor" strokeWidth={2.4} className="km-trip-flag"/><span>{trip.destination}</span></div></div>
                        <MiniMap geometry={trip.geometry}/>
                      </li>
                    ))}</ul>
                  )}
                </section>
              )}
            </div>
          </div>
        </div>

        {/* MARQUEE */}
        <div className="km-marquee"><div className="km-marquee-track"><span>{marqueeText}{marqueeText}{marqueeText}</span></div></div>
        <footer className="km-footer"><p>◊ DADOS LOCAIS · CRIPTOGRAFIA NATIVA ◊<br/><span className="km-coca">R$ 1,14 / KM</span> · ~ = NÚMERO APROXIMADO</p></footer>
      </main>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// CSS — RESPONSIVE BRUTALIST EDITORIAL
// Breakpoints: mobile default → 640px tablet → 1024px desktop → 1280px wide
// ═══════════════════════════════════════════════════════════════════════════
const CSS = `
/* ─── TOKENS ─── */
:root {
  --font-display: 'Geist', system-ui, sans-serif;
  --font-mono: 'Geist Mono', 'JetBrains Mono', ui-monospace, monospace;
  --font-serif: 'Instrument Serif', Georgia, serif;
  --font-body: 'Geist', system-ui, sans-serif;
  --ease-ios: cubic-bezier(0.32, 0.72, 0, 1);
  --ease-out-expo: cubic-bezier(0.16, 1, 0.3, 1);
  --container-max: 448px;
  --gutter: 16px;
  --radius-sm: 4px;
  --radius-md: 6px;
  --radius-lg: 8px;
  --shadow-sm: 0 1px 2px rgba(0, 0, 0, 0.05);
  --shadow-md: 0 2px 8px rgba(0, 0, 0, 0.08);
  --shadow-lg: 0 4px 16px rgba(0, 0, 0, 0.12);
  --focus-ring: 2px solid var(--coca-red);
  --transition-fast: 120ms var(--ease-ios);
  --transition-base: 180ms var(--ease-ios);
  --transition-slow: 240ms var(--ease-ios);
}
@media(min-width:640px){ :root { --container-max: 680px; --gutter: 24px; } }
@media(min-width:1024px){ :root { --container-max: 960px; --gutter: 32px; } }
@media(min-width:1280px){ :root { --container-max: 1120px; --gutter: 40px; } }

[data-theme="light"] {
  --bg-base:#FAFAFA;--bg-surface:#FFFFFF;--bg-elevated:#F5F5F5;--bg-overlay:#ECECEC;--bg-dark:#0A0A0A;
  --text-primary:#09090B;--text-secondary:#52525B;--text-muted:#71717A;
  --border-subtle:#D4D4D8;--border-default:#A1A1AA;--border-strong:#0A0A0A;
  --coca-red:#E61A27;--coca-red-hover:#D11620;--coca-red-active:#B80F1B;
  --coca-red-glow:rgba(230,26,39,0.18);--coca-red-soft:rgba(230,26,39,0.08);
  --status-green:#16A34A;--map-grid:rgba(10,9,8,0.06);
  --aurora:radial-gradient(ellipse 70% 60% at 50% -20%,rgba(230,26,39,0.08),transparent 70%);
}
[data-theme="dark"] {
  --bg-base:#0A0A0A;--bg-surface:#111111;--bg-elevated:#1A1A1A;--bg-overlay:#242424;--bg-dark:#050505;
  --text-primary:#EDEDED;--text-secondary:#A1A1AA;--text-muted:#71717A;
  --border-subtle:#1F1F1F;--border-default:#2A2A2A;--border-strong:#EDEDED;
  --coca-red:#FF4D5A;--coca-red-hover:#FF6670;--coca-red-active:#E63B47;
  --coca-red-glow:rgba(255,77,90,0.22);--coca-red-soft:rgba(255,77,90,0.10);
  --status-green:#4ADE80;--map-grid:rgba(255,255,255,0.06);
  --aurora:radial-gradient(ellipse 70% 60% at 50% -20%,rgba(255,77,90,0.12),transparent 70%);
}

*{box-sizing:border-box;}
*,.numeric{font-variant-numeric:tabular-nums lining-nums slashed-zero;font-feature-settings:'tnum' 1,'lnum' 1,'zero' 1;}
html,body{margin:0;padding:0;background:var(--bg-base);}
body{font-family:var(--font-body);color:var(--text-primary);-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;}

/* ─── ANIMATIONS ─── */
@keyframes fadeUp{from{opacity:0;transform:translateY(10px);}to{opacity:1;transform:translateY(0);}}
@keyframes slideDown{from{opacity:0;transform:translateY(-8px);}to{opacity:1;transform:translateY(0);}}
@keyframes fadeIn{from{opacity:0;}to{opacity:1;}}
@keyframes spin{from{transform:rotate(0);}to{transform:rotate(360deg);}}
@keyframes pulse{0%,49%{opacity:1;}50%,100%{opacity:0.35;}}
@keyframes marquee{from{transform:translateX(0);}to{transform:translateX(-33.33%);}}
@keyframes splashIn{from{opacity:0;transform:scale(0.95);}to{opacity:1;transform:scale(1);}}
@keyframes splashBar{0%{width:0;}100%{width:100%;}}
@keyframes routeDash{from{stroke-dashoffset:200;}to{stroke-dashoffset:0;}}
@keyframes loaderSlide{0%{transform:translateX(-100%);}100%{transform:translateX(100%);}}
@media(prefers-reduced-motion:reduce){*,*::before,*::after{animation-duration:0.01ms !important;transition-duration:0.01ms !important;}}
.km-fade-up{animation:fadeUp 0.4s var(--ease-ios) both;}
.km-fade-in{animation:fadeIn 0.3s var(--ease-ios) both;}
.km-spin{animation:spin 1s linear infinite;}

/* ─── APP SHELL ─── */
.km-app{background:var(--bg-base);color:var(--text-primary);min-height:100vh;font-family:var(--font-body);position:relative;overflow-x:hidden;}

/* ─── SPLASH ─── */
.km-splash{position:fixed;inset:0;z-index:9999;background:var(--coca-red);display:flex;align-items:center;justify-content:center;transition:opacity 0.55s var(--ease-ios),transform 0.55s var(--ease-ios);}
.km-splash--exit{opacity:0;transform:scale(1.04);pointer-events:none;}
.km-splash-grid{position:absolute;inset:0;background-image:linear-gradient(rgba(255,255,255,0.18) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,0.18) 1px,transparent 1px);background-size:32px 32px;opacity:0.4;}
.km-splash-content{position:relative;text-align:center;color:#FFF;animation:splashIn 0.5s var(--ease-out-expo) both;}
.km-splash-mono{font-family:var(--font-mono);font-size:11px;letter-spacing:0.32em;opacity:0.75;text-transform:uppercase;margin-bottom:18px;}
.km-splash-mono--small{margin-top:22px;opacity:0.6;}
.km-splash-title{font-family:var(--font-display);font-weight:900;font-size:64px;line-height:0.85;letter-spacing:-0.05em;display:flex;flex-direction:column;align-items:center;margin:0;}
.km-splash-italic{font-family:var(--font-serif);font-style:italic;font-weight:400;font-size:56px;margin-top:-4px;}
.km-splash-bar{width:180px;height:2px;background:rgba(255,255,255,0.2);margin:22px auto 0;overflow:hidden;}
.km-splash-bar-fill{height:100%;background:#FFF;animation:splashBar 1.1s var(--ease-out-expo) both;}
@media(min-width:640px){.km-splash-title{font-size:80px;}.km-splash-italic{font-size:72px;}.km-splash-bar{width:240px;}}
@media(min-width:1024px){.km-splash-title{font-size:100px;}.km-splash-italic{font-size:88px;}}

/* ─── TOP LOADER ─── */
.km-toploader{position:fixed;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,var(--coca-red),transparent);z-index:9998;animation:loaderSlide 1.2s linear infinite;}

/* ─── STATUS BAR ─── */
.km-statusbar{background:var(--bg-dark);border-bottom:1px solid var(--border-subtle);}
.km-statusbar-inner{max-width:var(--container-max);margin:0 auto;padding:6px var(--gutter);display:flex;justify-content:space-between;align-items:center;color:rgba(255,255,255,0.65);}
.km-status-dot{display:inline-block;width:6px;height:6px;background:var(--status-green);border-radius:50%;box-shadow:0 0 6px var(--status-green);animation:pulse 1.5s step-end infinite;margin-right:6px;vertical-align:middle;}
.km-status-dot--green{background:var(--status-green);}

/* ─── HEADER ─── */
.km-header{position:relative;background:var(--bg-base);border-bottom:2px solid var(--border-strong);padding:0 0 64px;overflow:hidden;}
.km-aurora-bg{position:absolute;inset:0;background:var(--aurora);pointer-events:none;}
.km-grid-overlay{position:absolute;inset:0;background-image:linear-gradient(var(--map-grid) 1px,transparent 1px),linear-gradient(90deg,var(--map-grid) 1px,transparent 1px);background-size:28px 28px;pointer-events:none;opacity:0.6;}
.km-header>*{position:relative;z-index:1;}
.km-header-inner{max-width:var(--container-max);margin:0 auto;padding:22px var(--gutter) 0;}
.km-header-top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:24px;}
.km-header-actions{display:flex;gap:6px;}
@media(min-width:640px){.km-header-actions{gap:8px;}}

.km-icon-btn{width:36px;height:36px;background:var(--bg-surface);border:1.5px solid var(--border-strong);border-radius:0;cursor:pointer;display:flex;align-items:center;justify-content:center;color:var(--text-primary);box-shadow:2px 2px 0 0 var(--border-strong);transition:transform 120ms var(--ease-ios),box-shadow 120ms var(--ease-ios);}
.km-icon-btn:active{transform:translate(2px,2px);box-shadow:0 0 0 0 var(--border-strong);}
@media(min-width:640px){.km-icon-btn{width:40px;height:40px;}}

.km-h1{font-family:var(--font-display);font-weight:900;font-size:clamp(44px,12vw,72px);line-height:0.85;letter-spacing:-0.045em;margin:8px 0 0;display:flex;flex-direction:column;}
.km-h1-italic{font-family:var(--font-serif);font-style:italic;font-weight:400;font-size:clamp(34px,9vw,56px);margin-top:-8px;color:var(--coca-red);}
@media(min-width:1024px){.km-h1{font-size:clamp(56px,5vw,80px);}.km-h1-italic{font-size:clamp(44px,4vw,64px);}}

.km-hero{margin-top:12px;}
.km-line{display:inline-block;width:18px;height:1.5px;background:var(--text-primary);margin-right:8px;vertical-align:middle;}
.km-hero-value{margin-top:6px;}
.km-hero-meta-row{margin-top:14px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;}
.km-pill{display:inline-flex;align-items:center;padding:5px 11px;border:1.5px solid var(--border-strong);background:var(--bg-surface);box-shadow:2px 2px 0 0 var(--border-strong);}
.km-wave{position:absolute;bottom:-1px;left:0;width:100%;height:56px;display:block;z-index:1;}

/* ─── BIG CURRENCY ─── */
.km-big-currency{display:inline-flex;align-items:baseline;font-family:var(--font-display);font-weight:900;letter-spacing:-0.05em;line-height:0.9;}
.km-big-currency-symbol{font-size:0.32em;font-weight:700;color:var(--text-secondary);margin-right:0.18em;align-self:flex-start;margin-top:0.18em;}
.km-big-currency-int{font-size:1em;color:var(--text-primary);}
.km-big-currency-dec{font-size:0.42em;font-weight:700;color:var(--text-secondary);margin-left:0.04em;}
.km-big-currency--hero{font-size:clamp(48px,13vw,96px);}
.km-big-currency--xl{font-size:clamp(40px,10vw,72px);}
.km-big-currency--md{font-size:clamp(32px,8vw,52px);}
@media(min-width:1024px){.km-big-currency--hero{font-size:clamp(64px,5vw,96px);}.km-big-currency--xl{font-size:clamp(52px,4vw,72px);}}

/* ─── MAIN ─── */
.km-main{max-width:var(--container-max);margin:-28px auto 0;padding:0 var(--gutter) 40px;position:relative;z-index:5;}

/* ─── CONTENT GRID (sidebar layout on desktop) ─── */
.km-content-grid{display:flex;flex-direction:column;gap:14px;}
@media(min-width:1024px){
  .km-content-grid{flex-direction:row;gap:24px;align-items:flex-start;}
  .km-content-left{width:380px;flex-shrink:0;position:sticky;top:16px;}
  .km-content-right{flex:1;min-width:0;}
}
@media(min-width:1280px){
  .km-content-left{width:420px;}
}

/* ─── ALERTS ─── */
.km-alert{display:flex;align-items:center;gap:10px;padding:12px 14px;border:1.5px solid var(--border-strong);border-radius:var(--radius-md);background:var(--bg-surface);margin-bottom:12px;box-shadow:var(--shadow-md);font-size:13px;font-weight:500;animation:slideDown 0.3s var(--ease-ios);}
.km-alert p{margin:0;flex:1;line-height:1.4;}
.km-alert button{background:none;border:none;cursor:pointer;color:var(--text-primary);padding:4px;border-radius:var(--radius-sm);transition:background-color var(--transition-fast);}
.km-alert button:hover{background-color:rgba(0,0,0,0.05);}
.km-alert button:focus-visible{outline:var(--focus-ring);outline-offset:2px;}
.km-alert--err{border-color:var(--coca-red);background:rgba(230,26,39,0.04);color:var(--coca-red);}
.km-alert--ok{border-color:var(--status-green);background:rgba(22,163,74,0.04);color:var(--status-green);}

/* ─── CARD ─── */
.km-card{background:var(--bg-surface);border:1.5px solid var(--border-strong);border-radius:var(--radius-lg);padding:18px;margin-bottom:16px;box-shadow:var(--shadow-md);position:relative;transition:box-shadow var(--transition-base),border-color var(--transition-base);}
.km-card:hover{box-shadow:var(--shadow-lg);}
@media(min-width:640px){.km-card{padding:24px;}}
@media(min-width:1024px){.km-card{padding:28px;}}
.km-card-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;}
.km-card-head--bordered{padding-bottom:12px;border-bottom:1.5px solid var(--border-default);}
.km-row-tight{display:flex;align-items:center;gap:8px;}
.km-redbar{display:inline-block;width:4px;height:22px;background:var(--coca-red);}
.km-h2{font-family:var(--font-display);font-weight:800;font-size:22px;letter-spacing:-0.035em;line-height:1;margin:0;}
@media(min-width:640px){.km-h2{font-size:26px;}}

/* ─── MONO / LABELS ─── */
.km-mono{font-family:var(--font-mono);font-weight:500;}
.km-mono-label{font-size:10.5px;letter-spacing:0.16em;text-transform:uppercase;color:var(--text-primary);font-weight:600;display:inline-flex;align-items:center;gap:6px;}
.km-mono-tiny{font-size:10px;letter-spacing:0.12em;text-transform:uppercase;}
.km-muted{color:var(--text-muted);}
.km-coca{color:var(--coca-red);font-weight:700;}
.km-italic{font-style:italic;}
.km-serif{font-family:var(--font-serif);}

/* ─── FIELDS ─── */
.km-field{margin-bottom:16px;}
.km-field-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;}
.km-textbtn{font-family:var(--font-mono);font-size:10px;letter-spacing:0.1em;color:var(--text-muted);background:none;border:none;cursor:pointer;padding:4px 6px;font-weight:500;transition:color var(--transition-fast);border-radius:var(--radius-sm);}
.km-textbtn:hover{color:var(--text-primary);}
.km-textbtn:focus-visible{outline:var(--focus-ring);outline-offset:2px;}
.km-textarea{width:100%;padding:11px 13px;font-size:14px;font-family:var(--font-body);color:var(--text-primary);background:var(--bg-base);border:1.5px solid var(--border-default);border-radius:var(--radius-md);resize:none;outline:none;line-height:1.45;font-weight:500;transition:border-color var(--transition-base),box-shadow var(--transition-base),background-color var(--transition-base);box-sizing:border-box;}
.km-textarea:hover{border-color:var(--coca-red-glow);background-color:var(--coca-red-soft);}
.km-textarea:focus{border-color:var(--coca-red);box-shadow:inset 0 0 0 1px var(--coca-red),0 0 0 3px var(--coca-red-soft);outline:none;}
.km-textarea:disabled{opacity:0.6;cursor:not-allowed;}
.km-autocomplete{position:relative;}
.km-suggestions{position:absolute;top:calc(100% + 4px);left:0;right:0;z-index:50;background:var(--bg-surface);border:1.5px solid var(--border-default);border-radius:var(--radius-md);list-style:none;margin:0;padding:4px 0;max-height:240px;overflow-y:auto;box-shadow:var(--shadow-lg);animation:fadeUp 0.2s var(--ease-ios);transform-origin:top;}
.km-suggestion-item{padding:10px 13px;font-size:13px;font-family:var(--font-body);color:var(--text-primary);cursor:pointer;border-bottom:1px solid var(--border-subtle);line-height:1.4;transition:background-color var(--transition-fast),color var(--transition-fast);}
.km-suggestion-item:last-child{border-bottom:none;}
.km-suggestion-item:hover{background:var(--coca-red-soft);color:var(--coca-red);font-weight:600;}
.km-suggestion-item:focus-visible{outline:var(--focus-ring);outline-offset:-2px;}
.km-dot{display:inline-block;width:9px;height:9px;flex-shrink:0;}
.km-dot--red{background:radial-gradient(circle at 30% 30%,var(--coca-red),var(--coca-red-active));border-radius:50%;box-shadow:0 0 0 3px var(--coca-red-soft);}

/* ─── BUTTONS ─── */
.km-btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;padding:11px 14px;border:1.5px solid var(--border-strong);border-radius:var(--radius-md);cursor:pointer;font-family:var(--font-display);font-weight:700;font-size:12px;letter-spacing:0.06em;text-transform:uppercase;background:var(--bg-surface);color:var(--text-primary);box-shadow:var(--shadow-sm);transition:transform var(--transition-fast),box-shadow var(--transition-fast),background var(--transition-fast),color var(--transition-fast);outline:none;}
.km-btn:hover:not(:disabled){transform:translateY(-2px);box-shadow:var(--shadow-md);}
.km-btn:focus-visible{outline:var(--focus-ring);outline-offset:2px;}
.km-btn:active:not(:disabled){transform:translateY(0px);box-shadow:var(--shadow-sm);}
.km-btn:disabled{opacity:0.45;cursor:not-allowed;}
.km-btn--block{width:100%;margin-top:8px;}
.km-btn--small{padding:8px 12px;font-size:11px;}
.km-btn--red{background:var(--coca-red);color:#FFF;border-color:var(--coca-red);}
.km-btn--red:hover:not(:disabled){background:var(--coca-red-hover);}
.km-btn--red:active:not(:disabled){background:var(--coca-red-active);}
.km-btn--ink{background:var(--text-primary);color:var(--bg-base);border-color:var(--text-primary);}
.km-btn--ink:hover:not(:disabled){opacity:0.9;}
.km-btn--save{background:var(--coca-red);color:#FFF;font-size:14px;padding:14px;margin-top:16px;border:2px solid var(--coca-red);border-radius:var(--radius-md);box-shadow:var(--shadow-md);letter-spacing:0.08em;font-weight:800;}
.km-btn--save:hover:not(:disabled){transform:translateY(-2px);box-shadow:var(--shadow-lg);}
.km-btn--save:focus-visible{outline:var(--focus-ring);outline-offset:2px;}
.km-btn--save:active:not(:disabled){transform:translateY(0px);}
.km-btn--press{cursor:pointer;}

/* ─── CONNECTOR ─── */
.km-connector{display:flex;align-items:center;gap:8px;padding:6px 0;color:var(--text-muted);}
.km-dashed{flex:1;height:1px;background:repeating-linear-gradient(90deg,var(--border-default) 0 4px,transparent 4px 8px);}

/* ─── ODOMETER ─── */
.km-odometer{background:var(--bg-dark);border:1.5px solid var(--border-strong);border-radius:var(--radius-lg);padding:16px 18px;margin:16px 0 16px;display:flex;align-items:center;justify-content:space-between;position:relative;overflow:hidden;box-shadow:var(--shadow-lg);transition:box-shadow var(--transition-base);}
.km-odometer:hover{box-shadow:var(--shadow-lg),inset 0 0 0 1px var(--coca-red-glow);}
.km-odometer-glow{position:absolute;top:-40px;right:-40px;width:140px;height:140px;border-radius:50%;background:radial-gradient(circle,var(--coca-red-glow),transparent 70%);pointer-events:none;}
.km-odometer-grid{position:absolute;inset:0;background-image:linear-gradient(rgba(255,255,255,0.06) 1px,transparent 1px);background-size:100% 12px;pointer-events:none;}
.km-odometer-info{position:relative;z-index:2;}
.km-odometer-label{color:rgba(255,255,255,0.6);letter-spacing:0.22em;}
.km-odometer-money{font-family:var(--font-mono);font-size:11px;color:var(--coca-red);font-weight:600;margin-top:4px;letter-spacing:0.04em;}
.km-odometer-sublabel{color:rgba(255,255,255,0.4);margin-top:3px;font-style:italic;}
.km-odometer-display{position:relative;z-index:2;}
.km-odometer-num{display:flex;align-items:baseline;gap:6px;font-family:var(--font-mono);font-weight:700;font-size:36px;color:#FFF;line-height:1;letter-spacing:-0.02em;text-shadow:0 0 18px var(--coca-red-glow);}
.km-odometer-unit{font-family:var(--font-display);font-size:14px;color:var(--coca-red);font-weight:800;letter-spacing:0.04em;font-style:italic;}
.km-odometer-empty{font-family:var(--font-mono);font-size:28px;color:rgba(255,255,255,0.22);font-weight:500;}
@media(min-width:640px){.km-odometer{padding:20px 24px;}.km-odometer-num{font-size:42px;}}

/* ─── TABS ─── */
.km-tabs{display:flex;background:var(--bg-surface);border:1.5px solid var(--border-default);border-radius:var(--radius-md);margin-bottom:16px;box-shadow:var(--shadow-sm);position:relative;z-index:1;padding:2px;}
.km-tab{flex:1;padding:10px 12px;background:transparent;border:none;border-radius:var(--radius-sm);cursor:pointer;font-family:var(--font-display);font-weight:700;font-size:11.5px;letter-spacing:0.08em;color:var(--text-muted);display:flex;align-items:center;justify-content:center;gap:6px;transition:all var(--transition-base);text-transform:uppercase;}
.km-tab:hover:not(.km-tab--active){background:var(--coca-red-soft);color:var(--text-primary);}
.km-tab:focus-visible{outline:var(--focus-ring);outline-offset:2px;}
.km-tab--active{background:var(--coca-red);color:#FFF;box-shadow:var(--shadow-sm);}
@media(min-width:640px){.km-tab{padding:12px 16px;font-size:12px;}}
.km-tabcontent{animation:fadeUp 0.35s var(--ease-ios) both;}

/* ─── HISTORY ─── */
.km-trip-list{list-style:none;padding:0;margin:0;}
.km-trip{padding:14px 0 10px;border-bottom:1px solid var(--border-default);}
.km-trip:first-child{padding-top:4px;}
.km-trip:last-child{border-bottom:2px solid var(--border-strong);padding-bottom:14px;}
.km-trip-head{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:9px;}
.km-trip-head-right{display:flex;align-items:center;gap:8px;}
.km-trip-km{display:flex;flex-direction:column;align-items:flex-end;gap:3px;}
.km-trip-km-pill{font-family:var(--font-mono);font-size:10.5px;font-weight:700;background:var(--coca-red);color:#FFF;padding:3px 9px;letter-spacing:0.04em;border:1.5px solid var(--border-strong);box-shadow:2px 2px 0 0 var(--border-strong);}
.km-trip-km-money{color:var(--status-green);font-weight:600;}
.km-icon-btn-tiny{background:none;border:none;cursor:pointer;color:var(--text-muted);padding:3px;display:flex;}
.km-icon-btn-tiny:hover{color:var(--coca-red);}
.km-trip-body{font-size:13px;line-height:1.5;color:var(--text-secondary);font-weight:500;}
.km-trip-line{display:flex;gap:10px;align-items:flex-start;}
.km-trip-line span:first-child{margin-top:5px;}
.km-trip-vline{margin-left:3.5px;height:11px;width:1px;background:var(--border-default);margin-top:1px;margin-bottom:1px;}
.km-trip-flag{margin-top:4px;flex-shrink:0;color:var(--text-primary);}

/* ─── MINI MAP ─── */
.km-minimap{position:relative;margin-top:10px;border:1.5px solid var(--border-strong);background:var(--bg-elevated);overflow:hidden;}
.km-minimap svg{display:block;width:100%;height:100px;}
@media(min-width:640px){.km-minimap svg{height:120px;}}
@media(min-width:1024px){.km-minimap svg{height:140px;}}
.km-minimap-label{position:absolute;top:6px;left:8px;right:8px;display:flex;justify-content:space-between;font-family:var(--font-mono);font-size:9px;letter-spacing:0.14em;color:var(--text-muted);text-transform:uppercase;pointer-events:none;}
.km-mini-empty{margin-top:10px;height:78px;background:repeating-linear-gradient(45deg,var(--bg-elevated),var(--bg-elevated) 8px,var(--bg-overlay) 8px,var(--bg-overlay) 16px);display:flex;align-items:center;justify-content:center;border:1.5px dashed var(--border-default);}
.km-mini-empty span{font-family:var(--font-mono);font-size:10px;color:var(--text-muted);letter-spacing:0.14em;text-transform:uppercase;}
.km-route-dash{stroke-dasharray:200;animation:routeDash 1.4s var(--ease-out-expo) forwards;}

/* ─── DASHBOARD GRIDS ─── */
.km-dash{display:block;}
.km-dash-top-grid{display:flex;flex-direction:column;gap:14px;}
.km-dash-mid-grid{display:flex;flex-direction:column;gap:14px;}
.km-dash-bottom-grid{display:flex;flex-direction:column;gap:14px;}

@media(min-width:640px){
  .km-dash-top-grid{flex-direction:row;gap:14px;}
  .km-dash-top-grid > *{flex:1;min-width:0;}
  .km-dash-mid-grid{flex-direction:row;gap:14px;align-items:stretch;}
  .km-dash-mid-grid .km-card--rings{flex:1.2;}
  .km-dash-mid-grid .km-stats-grid{flex:0.8;}
}

@media(min-width:1024px){
  .km-dash-bottom-grid{flex-direction:row;gap:14px;}
  .km-dash-bottom-grid > *{flex:1;min-width:0;}
}

/* HERO CARD */
.km-card--hero{position:relative;overflow:hidden;}
.km-aurora{position:absolute;top:-40px;right:-30px;width:200px;height:200px;background:radial-gradient(circle,var(--coca-red-glow),transparent 70%);pointer-events:none;}
.km-card--hero>*{position:relative;z-index:1;}
.km-hero-amount{margin:12px 0 14px;display:flex;letter-spacing:-0.05em;}
.km-hero-meta{display:flex;align-items:center;gap:10px;flex-wrap:wrap;font-size:11.5px;color:var(--text-secondary);}
.km-divider-vert{display:inline-block;width:1px;height:12px;background:var(--border-default);}
.km-hero-spark{margin-top:14px;display:flex;align-items:center;gap:10px;}
.km-sparkline{flex:1;height:32px;width:100%;}
@media(min-width:640px){.km-sparkline{height:40px;}}

.km-card--projection{background:var(--bg-elevated);}
.km-projection-content{margin-top:8px;}

/* RINGS */
.km-card--rings{padding-bottom:24px;}
.km-rings{display:flex;align-items:center;gap:18px;margin-top:14px;flex-wrap:wrap;justify-content:center;}
.km-rings-svg{width:160px;height:160px;flex-shrink:0;}
@media(min-width:640px){.km-rings-svg{width:180px;height:180px;}}
@media(min-width:1024px){.km-rings-svg{width:200px;height:200px;}}
.km-rings-legend{flex:1;min-width:120px;display:flex;flex-direction:column;gap:8px;}
.km-rings-legend-row{display:flex;align-items:center;gap:8px;font-family:var(--font-mono);font-size:11px;}
.km-rings-dot{width:8px;height:8px;border-radius:50%;}
.km-rings-legend-label{flex:1;letter-spacing:0.12em;color:var(--text-secondary);text-transform:uppercase;}
.km-rings-legend-pct{font-weight:700;color:var(--text-primary);letter-spacing:0.04em;}

/* STATS GRID */
.km-stats-grid{display:grid;grid-template-columns:1fr 1fr;gap:0;border:1.5px solid var(--border-strong);background:var(--bg-surface);margin-bottom:14px;box-shadow:4px 4px 0 0 var(--border-strong);}
@media(min-width:640px){.km-stats-grid{margin-bottom:0;align-self:stretch;}}
.km-stat-tile{padding:16px 14px;display:flex;flex-direction:column;gap:4px;border-right:1.5px solid var(--border-default);}
.km-stat-tile:last-child{border-right:none;}
@media(min-width:640px){
  .km-stats-grid{grid-template-columns:1fr;grid-template-rows:1fr 1fr;}
  .km-stat-tile{border-right:none;border-bottom:1.5px solid var(--border-default);}
  .km-stat-tile:last-child{border-bottom:none;}
}
@media(min-width:1024px){
  .km-stats-grid{grid-template-columns:1fr 1fr;grid-template-rows:auto;}
  .km-stat-tile{border-right:1.5px solid var(--border-default);border-bottom:none;}
  .km-stat-tile:last-child{border-right:none;}
}
.km-stat-num{font-family:var(--font-display);font-weight:900;font-size:clamp(28px,6vw,40px);letter-spacing:-0.04em;line-height:1;margin:4px 0 2px;}

/* PULL QUOTE */
.km-pullquote{margin:18px 0;padding-left:14px;border-left:2px solid var(--coca-red);}
@media(min-width:640px){.km-pullquote{padding-left:20px;}}
.km-pullquote p{font-size:clamp(18px,4vw,26px);line-height:1.2;margin:0 0 6px;font-weight:400;color:var(--text-primary);}
.km-pullquote-data{font-style:normal;font-weight:700;color:var(--coca-red);letter-spacing:0.02em;}
.km-pullquote cite{font-style:normal;color:var(--text-muted);}

/* ROWS (DAY / MONTH) */
.km-rows,.km-month-rows{list-style:none;padding:0;margin:0;}
.km-row{padding:12px 0 10px;border-bottom:1px solid var(--border-default);}
.km-row:first-child{padding-top:0;}.km-row:last-child{border-bottom:2px solid var(--border-strong);padding-bottom:12px;}
.km-row-main{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:7px;}
.km-row-date{font-size:12px;font-weight:700;letter-spacing:0.1em;}
.km-row-right{text-align:right;}
.km-row-km{font-size:13px;font-weight:700;letter-spacing:0.02em;}
.km-row-money{font-family:var(--font-display);font-style:italic;font-weight:800;font-size:13px;color:var(--coca-red);margin-top:2px;}
.km-bar{height:4px;background:var(--bg-elevated);position:relative;overflow:hidden;}
.km-bar-fill{height:100%;background:var(--coca-red);transition:width 0.6s var(--ease-out-expo);}
.km-month-row{display:flex;align-items:center;justify-content:space-between;padding:12px 0;border-bottom:1px solid var(--border-default);}
.km-month-row:first-child{padding-top:4px;}.km-month-row:last-child{border-bottom:2px solid var(--border-strong);}
.km-month-title{font-family:var(--font-display);font-weight:800;font-size:18px;letter-spacing:-0.02em;}
.km-month-money{font-family:var(--font-display);font-style:italic;font-weight:800;font-size:18px;color:var(--coca-red);letter-spacing:-0.01em;white-space:nowrap;}
@media(min-width:640px){.km-month-title{font-size:22px;}.km-month-money{font-size:22px;}}
.km-empty{font-size:12.5px;color:var(--text-muted);font-style:italic;text-align:center;padding:18px 0;margin:0;font-family:var(--font-serif);}
.km-pill-dark{background:var(--text-primary);color:var(--bg-base);padding:3px 8px;font-weight:600;}

/* ─── MARQUEE ─── */
.km-marquee{margin-top:28px;border-top:2px solid var(--border-strong);border-bottom:2px solid var(--border-strong);background:var(--bg-dark);color:#FFF;overflow:hidden;padding:8px 0;}
.km-marquee-track{display:flex;white-space:nowrap;animation:marquee 32s linear infinite;font-family:var(--font-mono);font-size:10.5px;letter-spacing:0.18em;text-transform:uppercase;font-weight:500;}
.km-marquee-track span{flex-shrink:0;color:rgba(255,255,255,0.85);}
@media(min-width:1024px){.km-marquee-track{font-size:11.5px;animation-duration:40s;}}
@media(prefers-reduced-motion:reduce){.km-marquee-track{animation:none;}}

/* ─── FOOTER ─── */
.km-footer{text-align:center;margin-top:22px;}
.km-footer p{margin:0;font-size:9.5px;line-height:1.7;letter-spacing:0.14em;text-transform:uppercase;font-family:var(--font-mono);color:var(--text-muted);font-weight:500;}
@media(min-width:640px){.km-footer p{font-size:10px;}}

/* ─── VIEW TRANSITIONS ─── */
@view-transition{navigation:auto;}
::view-transition-old(root),::view-transition-new(root){animation-duration:0.4s;animation-timing-function:cubic-bezier(0.32,0.72,0,1);}
`;
