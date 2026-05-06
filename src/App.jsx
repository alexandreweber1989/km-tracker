import { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { MapPin, Flag, Save, Download, Trash2, Loader2, AlertCircle, X, Navigation, ChevronDown, ChevronUp, ArrowDown, DollarSign, Calendar, TrendingUp, History, Sun, Moon, Volume2, VolumeX, Vibrate, Camera, FileText, ParkingCircle } from 'lucide-react';
import * as XLSX from 'xlsx';

const RATE = 1.14;
const APP_VERSION = 'v3·2'; // Manus Sync Active + Expenses

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

const TRIPS_KEY='km_trips_v1', EXPENSES_KEY='km_expenses_v1', THEME_KEY='coca_theme',SOUND_KEY='coca_sound',HAPTIC_KEY='coca_haptic';
const loadTrips=()=>{try{const r=localStorage.getItem(TRIPS_KEY);return r?JSON.parse(r):null;}catch{return null;}};
const saveTrips=(t)=>{try{localStorage.setItem(TRIPS_KEY,JSON.stringify(t));}catch{}};
const loadExpenses=()=>{try{const r=localStorage.getItem(EXPENSES_KEY);return r?JSON.parse(r):null;}catch{return null;}};
const saveExpenses=(e)=>{try{localStorage.setItem(EXPENSES_KEY,JSON.stringify(e));}catch{}};

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
  if(!values||values.length<2)return null;
  const max=Math.max(...values,0.01);
  const W=120,H=height;
  const points=values.map((v,i)=>`${(i/(values.length-1))*W},${H-(v/max)*H}`).join(' ');
  return(
    <svg width={W} height={H} viewBox={`0 0 ${W} ${H}`} style={{overflow:'visible'}}>
      <polyline points={points} fill="none" stroke="var(--coca-red)" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  );
}

// ═══ BIG CURRENCY ═══
function BigCurrency({value,size='md'}){
  const s=brlFmt.format(value);
  const symbol=s.slice(0,2);
  const rest=s.slice(2);
  const[integer,decimal]=rest.split(',');
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
      <div className="km-dash-top-grid">
        <section className="km-card km-card--hero">
          <div className="km-aurora"/>
          <header className="km-card-head"><span className="km-mono km-mono-label">[ TOTAL · ALL-TIME ]</span><span className="km-mono km-mono-label">{String(trips.length).padStart(3,'0')}/REC</span></header>
          <div className="km-hero-amount"><BigCurrency value={totalEarning} size="xl"/></div>
          <div className="km-hero-meta"><span className="km-mono">{numFmt.format(stats.totalKm)} KM</span><span className="km-divider-vert"/><span className="km-mono km-muted">× R$ 1,14/KM</span></div>
          {stats.last14.some(v=>v>0)&&(<div className="km-hero-spark"><Sparkline values={stats.last14} height={32}/><span className="km-mono km-mono-tiny">14d</span></div>)}
        </section>
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
      {stats.bestDay&&stats.totalKm>0&&(
        <blockquote className="km-pullquote"><p><span className="km-serif km-italic">Maior consumo da rota em </span><span className="km-mono km-pullquote-data">{stats.bestDay.date.slice(3)}</span><span className="km-serif km-italic">.</span></p><cite className="km-mono km-mono-tiny km-muted">— REGISTRO INTERNO · DRIVE LOG</cite></blockquote>
      )}
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
            <ul className="km-month-rows">{stats.months.map(m=>(
              <li key={m.key} className="km-month-row"><span className="km-month-title">{m.key}</span><span className="km-month-money">{brlFmt.format(m.km*RATE)}</span></li>
            ))}</ul>
          )}
        </section>
      </div>
    </div>
  );
}

// ═══ SPLASH ═══
function Splash({exiting}){
  return(
    <div className={`km-splash ${exiting?'km-splash--exit':''}`}>
      <div className="km-splash-inner">
        <div className="km-splash-logo"><span>DRIVE</span><span>LOG</span></div>
        <div className="km-splash-loader"><div className="km-splash-loader-bar"/></div>
        <div className="km-splash-meta">V3.2 · COCA-COLA BR · SYS ACTIVE</div>
      </div>
    </div>
  );
}

export default function KmTracker(){
  const[origin,setOrigin]=useState({address:'',lat:null,lng:null});
  const[destination,setDestination]=useState({address:'',lat:null,lng:null});
  const[distance,setDistance]=useState(null);
  const[distanceLabel,setDistanceLabel]=useState('');
  const[routeGeometry,setRouteGeometry]=useState(null);
  const[trips,setTrips]=useState([]);
  const[expenses,setExpenses]=useState([]);
  const[loading,setLoading]=useState({origin:false,destination:false,distance:false,ocr:false});
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
  const fileInputRef=useRef(null);
  const feedback=useFeedback();

  useEffect(()=>{setThemeState(getInitialTheme());setSoundOn(localStorage.getItem(SOUND_KEY)!=='off');setHapticOn(localStorage.getItem(HAPTIC_KEY)!=='off');const t1=setTimeout(()=>setSplashExiting(true),1100);const t2=setTimeout(()=>setSplashGone(true),1700);return()=>{clearTimeout(t1);clearTimeout(t2);};},[]);
  useEffect(()=>{const update=()=>{const d=new Date();setTime(`${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`);};update();const id=setInterval(update,1000);return()=>clearInterval(id);},[]);
  useEffect(()=>{
    const loadedTrips=loadTrips();if(loadedTrips&&Array.isArray(loadedTrips))setTrips(loadedTrips);else{setTrips(SEED_TRIPS);saveTrips(SEED_TRIPS);}
    const loadedExpenses=loadExpenses();if(loadedExpenses&&Array.isArray(loadedExpenses))setExpenses(loadedExpenses);
  },[]);
  useEffect(()=>{if(trips.length>0)saveTrips(trips);},[trips]);
  useEffect(()=>{if(expenses.length>0)saveExpenses(expenses);},[expenses]);
  useEffect(()=>{if(initialized.current)return;initialized.current=true;const fonts=document.createElement('link');fonts.href='https://fonts.googleapis.com/css2?family=Geist:wght@300;400;500;600;700;800;900&family=Geist+Mono:wght@400;500;700&family=Instrument+Serif:ital@0;1&display=swap';fonts.rel='stylesheet';document.head.appendChild(fonts);const style=document.createElement('style');style.textContent=CSS;document.head.appendChild(style);},[]);

  useEffect(()=>{if(origin.lat==null||destination.lat==null){setDistance(null);setDistanceLabel('');setRouteGeometry(null);return;}let cancelled=false;(async()=>{setLoading(s=>({...s,distance:true}));try{const r=await fetch(`https://router.project-osrm.org/route/v1/driving/${origin.lng},${origin.lat};${destination.lng},${destination.lat}?overview=simplified&geometries=geojson`);const data=await r.json();if(cancelled)return;if(data.routes?.length>0){const route=data.routes[0];setDistance(route.distance/1000);setDistanceLabel('rota de carro');setRouteGeometry(route.geometry?.coordinates||null);}else throw new Error();}catch{if(cancelled)return;setDistance(haversineKm(origin.lat,origin.lng,destination.lat,destination.lng));setDistanceLabel('linha reta · rota indisponível');setRouteGeometry([[origin.lng,origin.lat],[destination.lng,destination.lat]]);}finally{if(!cancelled)setLoading(s=>({...s,distance:false}));}})();return()=>{cancelled=true;};},[origin.lat,origin.lng,destination.lat,destination.lng]);

  async function searchAddress(query,which){if(query.length<3){if(which==='origin'){setOriginSuggestions([]);setShowOriginDrop(false);}else{setDestSuggestions([]);setShowDestDrop(false);}return;}const typedNum=query.match(/[\s,]+(\d{1,5})(?:\s*[-,]|\s*$)/);const num=typedNum?typedNum[1]:null;try{const url=`https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(query)}&format=json&addressdetails=1&accept-language=pt-BR&countrycodes=br&limit=5`;const r=await fetch(url);if(!r.ok)return;const list=await r.json();const results=list.map(item=>{const a=item.address||{};const street=a.road||a.pedestrian||a.path||'';const houseNum=a.house_number||num||'';const neighborhood=a.suburb||a.neighbourhood||a.quarter||a.city_district||'';const city=a.city||a.town||a.village||a.municipality||'';const stateAbbr=STATE_MAP[a.state||'']||a.state||'';let label=street;if(houseNum)label+=`, ${houseNum}`;if(neighborhood)label+=` - ${neighborhood}`;if(city)label+=`, ${city}`;if(stateAbbr)label+=` - ${stateAbbr}`;return{label:label||item.display_name,lat:parseFloat(item.lat),lng:parseFloat(item.lon)};});if(which==='origin'){setOriginSuggestions(results);setShowOriginDrop(results.length>0);}else{setDestSuggestions(results);setShowDestDrop(results.length>0);}}catch{}}
  function handleAddressInput(e,which){const val=e.target.value;const setter=which==='origin'?setOrigin:setDestination;setter(prev=>({...prev,address:val,lat:null,lng:null}));const debRef=which==='origin'?originDebounce:destDebounce;clearTimeout(debRef.current);debRef.current=setTimeout(()=>searchAddress(val,which),400);}
  function selectSuggestion(item,which){const setter=which==='origin'?setOrigin:setDestination;setter({address:item.label,lat:item.lat,lng:item.lng});if(which==='origin'){setOriginSuggestions([]);setShowOriginDrop(false);}else{setDestSuggestions([]);setShowDestDrop(false);}feedback('tap');}
  async function capture(which){setError(null);setLoading(s=>({...s,[which]:true}));setTopLoading(true);feedback('tap');try{if(!navigator.geolocation)throw new Error('Geolocalização não disponível');const pos=await new Promise((res,rej)=>navigator.geolocation.getCurrentPosition(p=>res({lat:p.coords.latitude,lng:p.coords.longitude}),e=>rej(new Error(e.code===1?'Permissão negada. Habilite o GPS.':e.code===2?'GPS indisponível.':e.code===3?'Timeout. Tente novamente.':'Erro de localização')),{enableHighAccuracy:true,timeout:20000,maximumAge:0}));const result=await getAddressWithNumber(pos.lat,pos.lng);const address=formatAddr(result.data,result.number,result.isApprox);(which==='origin'?setOrigin:setDestination)({address,lat:pos.lat,lng:pos.lng});feedback('success');if(result.number&&result.isApprox){setSuccess('Número aproximado · pode ajustar');setTimeout(()=>setSuccess(null),2600);}else if(!result.number){setSuccess('Sem número · adicione manual');setTimeout(()=>setSuccess(null),2600);}}catch(e){setError(e.message);feedback('error');}finally{setLoading(s=>({...s,[which]:false}));setTopLoading(false);}}
  function clearLocation(w){feedback('tap');(w==='origin'?setOrigin:setDestination)({address:'',lat:null,lng:null});if(w==='origin'){setOriginSuggestions([]);setShowOriginDrop(false);}else{setDestSuggestions([]);setShowDestDrop(false);}}
  function handleSave(){if(!origin.address.trim()||!destination.address.trim()){setError('Preencha origem e destino');feedback('error');return;}const d=new Date();const newTrip={id:'t'+Date.now(),date:`${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`,origin:origin.address.trim(),destination:destination.address.trim(),km:distance!=null?Number(distance.toFixed(2)):null,kmLabel:distanceLabel,geometry:routeGeometry};setTrips(prev=>[newTrip,...prev]);setOrigin({address:'',lat:null,lng:null});setDestination({address:'',lat:null,lng:null});setDistance(null);setDistanceLabel('');setRouteGeometry(null);setError(null);setSuccess('Viagem registrada');setTimeout(()=>setSuccess(null),2200);feedback('success');}
  function deleteTrip(id){if(window.confirm('Remover esta viagem?')){setTrips(prev=>prev.filter(t=>t.id!==id));feedback('swoosh');}}
  function deleteExpense(id){if(window.confirm('Remover esta despesa?')){setExpenses(prev=>prev.filter(e=>e.id!==id));feedback('swoosh');}}

  async function handleFileUpload(e){
    const file=e.target.files[0];if(!file)return;
    setLoading(s=>({...s,ocr:true}));setTopLoading(true);feedback('tap');
    const formData=new FormData();formData.append('receipt',file);
    try{
      const r=await fetch('http://localhost:3001/api/process-receipt',{method:'POST',body:formData});
      const data=await r.json();
      if(data.error)throw new Error(data.error);
      const newExpense={id:'e'+Date.now(),...data,date:data.data||new Date().toLocaleDateString('pt-BR')};
      setExpenses(prev=>[newExpense,...prev]);
      setSuccess(`${data.tipo==='pedagio'?'Pedágio':'Estacionamento'} processado`);
      setTimeout(()=>setSuccess(null),2200);feedback('success');
    }catch(err){setError(err.message);feedback('error');}finally{setLoading(s=>({...s,ocr:false}));setTopLoading(false);e.target.value='';}
  }

  function exportToExcel(){
    if(!trips.length&&!expenses.length){setError('Sem dados para exportar');feedback('error');return;}
    const wb=XLSX.utils.book_new();
    if(trips.length){
      const sortedTrips=[...trips].sort((a,b)=>{const p=s=>{const[d,m,y]=s.split('/');return new Date(y,m-1,d);};return p(a.date)-p(b.date);});
      const tripData=[['Data','Origem','Destino','KM','Valor'],...sortedTrips.map(t=>[t.date,t.origin,t.destination,t.km,t.km?t.km*RATE:0])];
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(tripData),'Viagens');
    }
    if(expenses.length){
      const sortedExpenses=[...expenses].sort((a,b)=>{const p=s=>{const[d,m,y]=s.split('/');return new Date(y,m-1,d);};return p(a.date)-p(b.date);});
      const expData=[['Data','Tipo','Local/Concessionária','Placa','Valor'],...sortedExpenses.map(e=>[e.date,e.tipo,e.concessionaria||e.local||'',e.placa||'',e.valor])];
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(expData),'Despesas');
    }
    XLSX.writeFile(wb,'controle_drive_log.xlsx');
    setSuccess('Planilha exportada');setTimeout(()=>setSuccess(null),2200);feedback('success');
  }

  function handleToggleTheme(){feedback('tap');setThemeState(toggleTheme(theme));}
  function handleToggleSound(){const next=!soundOn;setSoundOn(next);localStorage.setItem(SOUND_KEY,next?'on':'off');if(next)feedback('tap');}
  function handleToggleHaptic(){const next=!hapticOn;setHapticOn(next);localStorage.setItem(HAPTIC_KEY,next?'on':'off');if(next&&'vibrate'in navigator)navigator.vibrate(15);}
  function handleTabSwitch(t){if(t===activeTab)return;feedback('swoosh');if(document.startViewTransition)document.startViewTransition(()=>setActiveTab(t));else setActiveTab(t);}

  const totalKm=trips.reduce((s,t)=>s+(t.km||0),0);
  const totalExpenses=expenses.reduce((s,e)=>s+(Number(e.valor)||0),0);
  const totalEarning=totalKm*RATE+totalExpenses;
  const tripsWithKm=trips.filter(t=>t.km!=null).length;
  const canSave=origin.address.trim()&&destination.address.trim();
  const marqueeText=' · DRIVE LOG · COCA-COLA BR · FROTA · '+numFmt.format(totalKm)+' KM · '+brlFmt.format(totalEarning)+' · '+String(trips.length).padStart(3,'0')+' VIAGENS · '+brlFmt.format(totalExpenses)+' DESPESAS';

  return(
    <div className="km-app" data-theme={theme}>
      {!splashGone&&<Splash exiting={splashExiting}/>}
      {topLoading&&<div className="km-toploader"/>}
      <div className="km-statusbar"><div className="km-statusbar-inner">
        <span className="km-mono km-mono-tiny"><span className="km-status-dot"/> SYS · ONLINE</span>
        <span className="km-mono km-mono-tiny km-muted">{APP_VERSION} · DRIVE LOG</span>
        <span className="km-mono km-mono-tiny">{time}</span>
      </div></div>
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
          <div className="km-hero">
            <div className="km-mono km-mono-label"><span className="km-line"/>VALOR A RECEBER</div>
            <div className="km-hero-value"><BigCurrency value={totalEarning} size="hero"/></div>
            <div className="km-hero-meta-row">
              <span className="km-pill"><span className="km-status-dot km-status-dot--green"/><span className="km-mono km-mono-tiny">{numFmt.format(totalKm)} KM TOTAIS</span></span>
              <span className="km-pill"><span className="km-status-dot" style={{background:'var(--text-primary)'}}/><span className="km-mono km-mono-tiny">{brlFmt.format(totalExpenses)} DESPESAS</span></span>
            </div>
          </div>
        </div>
        <svg viewBox="0 0 400 60" preserveAspectRatio="none" className="km-wave"><path d="M0,40 C60,10 120,55 200,30 C280,5 340,50 400,25 L400,60 L0,60 Z" fill="var(--bg-base)"/><path d="M0,40 C60,10 120,55 200,30 C280,5 340,50 400,25" fill="none" stroke="rgba(255,255,255,0.4)" strokeWidth="1"/></svg>
      </header>
      <main className="km-main">
        <div className="km-content-grid">
          <div className="km-content-left">
            {error&&(<div className="km-alert km-alert--err km-fade-in"><AlertCircle size={18} strokeWidth={2.4}/><p>{error}</p><button onClick={()=>setError(null)}><X size={14}/></button></div>)}
            {success&&(<div className="km-alert km-alert--ok km-fade-in"><span className="km-status-dot km-status-dot--green"/><p>{success}</p></div>)}

            <section className="km-card">
              <header className="km-card-head km-card-head--bordered"><div className="km-row-tight"><span className="km-redbar"/><h2 className="km-h2">Nova viagem</h2></div><span className="km-mono km-mono-tiny km-muted">001/REC</span></header>
              <div className="km-field"><div className="km-field-head"><label htmlFor="origin-address" className="km-mono km-mono-label"><span className="km-dot km-dot--red"/>· ORIGEM</label>{origin.address&&<button onClick={()=>clearLocation('origin')} className="km-textbtn" aria-label="Limpar endereço de origem">CLEAR</button>}</div><div className="km-autocomplete"><textarea id="origin-address" value={origin.address} onChange={e=>handleAddressInput(e,'origin')} onFocus={()=>originSuggestions.length>0&&setShowOriginDrop(true)} onBlur={()=>setTimeout(()=>setShowOriginDrop(false),200)} placeholder="Digite o endereço de partida" rows={2} className="km-textarea" aria-label="Endereço de origem" aria-describedby="origin-help" aria-autocomplete="list" aria-controls="origin-suggestions"/>{showOriginDrop&&originSuggestions.length>0&&(<ul id="origin-suggestions" className="km-suggestions" role="listbox">{originSuggestions.map((s,i)=>(<li key={i} onMouseDown={()=>selectSuggestion(s,'origin')} className="km-suggestion-item" role="option">{s.label}</li>))}</ul>)}</div><div id="origin-help" className="km-mono km-mono-tiny km-muted" style={{marginTop:'4px'}}>Comece a digitar para autocomplete ou clique abaixo para usar GPS</div><button onClick={()=>capture('origin')} disabled={loading.origin} className="km-btn km-btn--ink km-btn--block km-press" aria-label={loading.origin?'Buscando localização de origem':'Capturar localização de origem'}>{loading.origin?<><Loader2 size={15} className="km-spin"/> BUSCANDO…</>:<><MapPin size={15} strokeWidth={2.4}/> CAPTURAR LOCALIZAÇÃO</>}</button></div>
              <div className="km-dashed" style={{margin:'20px 0'}}/>
              <div className="km-field"><div className="km-field-head"><label htmlFor="dest-address" className="km-mono km-mono-label"><span className="km-dot km-dot--red"/>· DESTINO</label>{destination.address&&<button onClick={()=>clearLocation('destination')} className="km-textbtn" aria-label="Limpar endereço de destino">CLEAR</button>}</div><div className="km-autocomplete"><textarea id="dest-address" value={destination.address} onChange={e=>handleAddressInput(e,'destination')} onFocus={()=>destSuggestions.length>0&&setShowDestDrop(true)} onBlur={()=>setTimeout(()=>setShowDestDrop(false),200)} placeholder="Digite o endereço de chegada" rows={2} className="km-textarea" aria-label="Endereço de destino" aria-describedby="dest-help" aria-autocomplete="list" aria-controls="dest-suggestions"/>{showDestDrop&&destSuggestions.length>0&&(<ul id="dest-suggestions" className="km-suggestions" role="listbox">{destSuggestions.map((s,i)=>(<li key={i} onMouseDown={()=>selectSuggestion(s,'destination')} className="km-suggestion-item" role="option">{s.label}</li>))}</ul>)}</div><div id="dest-help" className="km-mono km-mono-tiny km-muted" style={{marginTop:'4px'}}>Comece a digitar para autocomplete ou clique abaixo para usar GPS</div><button onClick={()=>capture('destination')} disabled={loading.destination} className="km-btn km-btn--ink km-btn--block km-press" aria-label={loading.destination?'Buscando localização de destino':'Capturar localização de destino'}>{loading.destination?<><Loader2 size={15} className="km-spin"/> BUSCANDO…</>:<><Flag size={15} strokeWidth={2.4}/> CAPTURAR LOCALIZAÇÃO</>}</button></div>
              <div className="km-odometer" role="region" aria-live="polite" aria-label="Informações de distância"><div className="km-odometer-glow"/><div className="km-odometer-grid"/><div className="km-odometer-info"><div className="km-mono km-mono-label km-odometer-label">◊ DISTÂNCIA</div>{distance!=null&&<div className="km-odometer-money" aria-label={`Valor a receber: ${brlFmt.format(distance*RATE)}`}>{brlFmt.format(distance*RATE)}</div>}{distanceLabel&&<div className="km-mono km-mono-tiny km-odometer-sublabel">{distanceLabel}</div>}</div><div className="km-odometer-display">{loading.distance?<Loader2 size={28} className="km-spin" style={{color:'var(--coca-red)'}} aria-label="Calculando distância"/>:distance!=null?<div className="km-odometer-num"><span aria-label={`${numFmt.format(distance)} quilômetros`}>{numFmt.format(distance)}</span><span className="km-odometer-unit">KM</span></div>:<div className="km-odometer-empty">--.--</div>}</div></div>
              <button onClick={handleSave} disabled={!canSave} className="km-btn km-btn--save km-btn--block km-press" aria-label={!canSave?'Preencha origem e destino para registrar viagem':'Registrar viagem'}><Save size={16} strokeWidth={2.4}/> REGISTRAR VIAGEM</button>
            </section>

            <section className="km-card">
              <header className="km-card-head km-card-head--bordered"><div className="km-row-tight"><span className="km-redbar"/><h2 className="km-h2">Despesas</h2></div><span className="km-mono km-mono-tiny km-muted">AUTO-SCAN</span></header>
              <p className="km-mono km-mono-tiny km-muted" style={{marginBottom:12}}>Tire uma foto do recibo de pedágio ou estacionamento para processamento automático.</p>
              <input type="file" accept="image/*" capture="environment" ref={fileInputRef} onChange={handleFileUpload} style={{display:'none'}}/>
              <button onClick={()=>fileInputRef.current.click()} disabled={loading.ocr} className="km-btn km-btn--red km-btn--block km-press" style={{padding:'16px'}}>
                {loading.ocr?<><Loader2 size={18} className="km-spin"/> PROCESSANDO…</>:<><Camera size={18} strokeWidth={2.4}/> ESCANEAR RECIBO</>}
              </button>
            </section>
          </div>
          <div className="km-content-right">
            <div className="km-tabs">
              <button onClick={()=>handleTabSwitch('history')} className={`km-tab ${activeTab==='history'?'km-tab--active':''}`}><History size={14} strokeWidth={2.4}/> VIAGENS</button>
              <button onClick={()=>handleTabSwitch('expenses')} className={`km-tab ${activeTab==='expenses'?'km-tab--active':''}`}><FileText size={14} strokeWidth={2.4}/> DESPESAS</button>
              <button onClick={()=>handleTabSwitch('dashboard')} className={`km-tab ${activeTab==='dashboard'?'km-tab--active':''}`}><TrendingUp size={14} strokeWidth={2.4}/> DASHBOARD</button>
            </div>
            <div className="km-tabcontent" key={activeTab}>
              {activeTab==='dashboard'?(<Dashboard trips={trips}/>):activeTab==='expenses'?(
                <section className="km-card km-fade-up">
                  <header className="km-card-head km-card-head--bordered"><div className="km-row-tight"><span className="km-redbar"/><h2 className="km-h2">Recibos</h2><span className="km-mono km-mono-tiny km-pill-dark">{String(expenses.length).padStart(2,'0')}</span></div><button onClick={exportToExcel} className="km-btn km-btn--ink km-btn--small km-press"><Download size={13} strokeWidth={2.4}/> EXPORTAR</button></header>
                  {expenses.length===0?(<p className="km-empty">Nenhum recibo escanerado ainda.</p>):(
                    <ul className="km-trip-list">{expenses.map((exp)=>(
                      <li key={exp.id} className="km-trip">
                        <div className="km-trip-head">
                          <span className="km-mono km-mono-tiny">{exp.date} <span className="km-muted">·</span> {exp.tipo==='pedagio'?'PEDÁGIO':'ESTACIONAMENTO'}</span>
                          <div className="km-trip-head-right">
                            <div className="km-trip-km"><span className="km-trip-km-pill" style={{background:'var(--text-primary)'}}>{brlFmt.format(exp.valor)}</span></div>
                            <button onClick={()=>deleteExpense(exp.id)} className="km-icon-btn-tiny"><Trash2 size={13}/></button>
                          </div>
                        </div>
                        <div className="km-trip-body">
                          <div className="km-trip-line">{exp.tipo==='pedagio'?<Navigation size={11} strokeWidth={2.4} className="km-trip-flag"/>:<ParkingCircle size={11} strokeWidth={2.4} className="km-trip-flag"/>}<span>{exp.concessionaria||exp.local}</span></div>
                          {exp.placa&&<div className="km-trip-line" style={{marginTop:4,fontSize:11}}><span className="km-mono km-muted">PLACA: {exp.placa}</span></div>}
                        </div>
                        {exp.processed_image && (
                          <div className="km-minimap" style={{marginTop:12,height:'auto'}}>
                            <img src={`http://localhost:3001${exp.processed_image}`} alt="Recibo" style={{width:'100%',display:'block',borderRadius:'4px'}}/>
                            <div className="km-minimap-label"><span>◊ SCANNER ATIVO</span><a href={`http://localhost:3001${exp.processed_image}`} download className="km-mono" style={{color:'var(--coca-red)',textDecoration:'none'}}>BAIXAR</a></div>
                          </div>
                        )}
                      </li>
                    ))}</ul>
                  )}
                </section>
              ):(
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
        <div className="km-marquee"><div className="km-marquee-track"><span>{marqueeText}{marqueeText}{marqueeText}</span></div></div>
        <footer className="km-footer"><p>◊ DADOS LOCAIS · CRIPTOGRAFIA NATIVA ◊<br/><span className="km-coca">R$ 1,14 / KM</span> · ~ = NÚMERO APROXIMADO</p></footer>
      </main>
    </div>
  );
}

const CSS = `
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
  --coca-red: #E61A27;
  --coca-red-hover: #CC1723;
  --coca-red-active: #B3141F;
  --coca-red-soft: rgba(230,26,39,0.08);
  --coca-red-glow: rgba(230,26,39,0.25);
  --status-green: #16A34A;
}
@media(min-width:640px){ :root { --container-max: 680px; --gutter: 24px; } }
@media(min-width:1024px){ :root { --container-max: 960px; --gutter: 32px; } }
@media(min-width:1280px){ :root { --container-max: 1120px; --gutter: 40px; } }
[data-theme="light"] {
  --bg-base:#FAFAFA;--bg-surface:#FFFFFF;--bg-elevated:#F5F5F5;--bg-overlay:#ECECEC;--bg-dark:#0A0A0A;
  --text-primary:#0A0A0A;--text-secondary:#404040;--text-muted:#737373;
  --border-default:#E5E5E5;--border-strong:#D4D4D4;--border-subtle:#F0F0F0;
  --map-grid: rgba(0,0,0,0.05);
}
[data-theme="dark"] {
  --bg-base:#0A0A0A;--bg-surface:#121212;--bg-elevated:#1A1A1A;--bg-overlay:#262626;--bg-dark:#000000;
  --text-primary:#F5F5F5;--text-secondary:#A3A3A3;--text-muted:#737373;
  --border-default:#262626;--border-strong:#404040;--border-subtle:#1A1A1A;
  --map-grid: rgba(255,255,255,0.05);
}
body{margin:0;padding:0;background:var(--bg-base);color:var(--text-primary);font-family:var(--font-body);-webkit-font-smoothing:antialiased;overflow-x:hidden;}
.km-app{min-height:100vh;display:flex;flex-direction:column;background:var(--bg-base);transition:background-color var(--transition-base),color var(--transition-base);}
.km-toploader{position:fixed;top:0;left:0;right:0;height:3px;background:var(--coca-red);z-index:2000;animation:loadingBar 2s infinite ease-in-out;}
@keyframes loadingBar{0%{transform:translateX(-100%);}100%{transform:translateX(100%);}}
.km-statusbar{background:var(--bg-dark);color:#FFF;padding:6px 0;border-bottom:1px solid var(--border-strong);position:relative;z-index:100;}
.km-statusbar-inner{max-width:var(--container-max);margin:0 auto;padding:0 var(--gutter);display:flex;justify-content:space-between;align-items:center;}
.km-status-dot{width:6px;height:6px;background:var(--status-green);border-radius:50%;display:inline-block;margin-right:4px;box-shadow:0 0 8px var(--status-green);}
.km-status-dot--green{background:var(--status-green);box-shadow:0 0 8px var(--status-green);}
.km-header{background:var(--bg-dark);color:#FFF;padding:40px 0 60px;position:relative;overflow:hidden;}
.km-header-inner{max-width:var(--container-max);margin:0 auto;padding:0 var(--gutter);position:relative;z-index:2;}
.km-header-top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:32px;}
.km-header-actions{display:flex;gap:12px;}
.km-h1{font-family:var(--font-display);font-weight:900;font-size:38px;letter-spacing:-0.05em;line-height:0.85;margin:8px 0 0;text-transform:uppercase;}
.km-h1 span{display:block;}
.km-h1-italic{font-family:var(--font-serif);font-style:italic;font-weight:400;text-transform:lowercase;letter-spacing:-0.02em;color:var(--coca-red);margin-left:4px;}
.km-hero{margin-top:24px;}
.km-hero-value{margin:8px 0 12px;}
.km-hero-meta-row{display:flex;align-items:center;gap:12px;flex-wrap:wrap;}
.km-pill{background:rgba(255,255,255,0.1);padding:4px 10px;border-radius:100px;display:inline-flex;align-items:center;gap:6px;border:1px solid rgba(255,255,255,0.1);}
.km-big-currency{font-family:var(--font-display);font-weight:900;line-height:1;display:inline-flex;align-items:baseline;letter-spacing:-0.04em;}
.km-big-currency--hero{font-size:72px;}
.km-big-currency--xl{font-size:48px;}
.km-big-currency--md{font-size:32px;}
.km-big-currency-symbol{font-size:0.4em;margin-right:4px;color:var(--coca-red);}
.km-big-currency-dec{font-size:0.5em;opacity:0.6;margin-left:2px;}
.km-wave{position:absolute;bottom:0;left:0;width:100%;height:60px;z-index:1;}
.km-main{max-width:var(--container-max);margin:-30px auto 40px;padding:0 var(--gutter);position:relative;z-index:10;}
.km-content-grid{display:flex;flex-direction:column;gap:14px;}
@media(min-width:1024px){
  .km-content-grid{flex-direction:row;gap:24px;align-items:flex-start;}
  .km-content-left{width:380px;flex-shrink:0;position:sticky;top:16px;}
  .km-content-right{flex:1;min-width:0;}
}
.km-card{background:var(--bg-surface);border:1.5px solid var(--border-strong);border-radius:var(--radius-lg);padding:18px;margin-bottom:16px;box-shadow:var(--shadow-md);position:relative;}
.km-card-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;}
.km-card-head--bordered{padding-bottom:12px;border-bottom:1.5px solid var(--border-default);}
.km-h2{font-family:var(--font-display);font-weight:800;font-size:22px;letter-spacing:-0.035em;line-height:1;margin:0;}
.km-redbar{display:inline-block;width:4px;height:22px;background:var(--coca-red);}
.km-row-tight{display:flex;align-items:center;gap:8px;}
.km-mono{font-family:var(--font-mono);font-weight:500;}
.km-mono-label{font-size:10.5px;letter-spacing:0.16em;text-transform:uppercase;color:var(--text-primary);font-weight:600;display:inline-flex;align-items:center;gap:6px;}
.km-mono-tiny{font-size:10px;letter-spacing:0.12em;text-transform:uppercase;}
.km-muted{color:var(--text-muted);}
.km-btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;padding:11px 14px;border:1.5px solid var(--border-strong);border-radius:var(--radius-md);cursor:pointer;font-family:var(--font-display);font-weight:700;font-size:12px;letter-spacing:0.06em;text-transform:uppercase;background:var(--bg-surface);color:var(--text-primary);transition:all var(--transition-fast);}
.km-btn--block{width:100%;}
.km-btn--red{background:var(--coca-red);color:#FFF;border-color:var(--coca-red);}
.km-btn--ink{background:var(--text-primary);color:var(--bg-base);border-color:var(--text-primary);}
.km-btn--save{background:var(--coca-red);color:#FFF;font-size:14px;padding:14px;margin-top:16px;border:2px solid var(--coca-red);font-weight:800;}
.km-textarea{width:100%;padding:11px 13px;font-size:14px;font-family:var(--font-body);color:var(--text-primary);background:var(--bg-base);border:1.5px solid var(--border-default);border-radius:var(--radius-md);resize:none;outline:none;line-height:1.45;box-sizing:border-box;}
.km-tabs{display:flex;background:var(--bg-surface);border:1.5px solid var(--border-default);border-radius:var(--radius-md);margin-bottom:16px;padding:2px;}
.km-tab{flex:1;padding:10px 12px;background:transparent;border:none;border-radius:var(--radius-sm);cursor:pointer;font-family:var(--font-display);font-weight:700;font-size:11px;color:var(--text-muted);display:flex;align-items:center;justify-content:center;gap:6px;text-transform:uppercase;transition:all var(--transition-base);}
.km-tab--active{background:var(--coca-red);color:#FFF;}
.km-trip-list{list-style:none;padding:0;margin:0;}
.km-trip{padding:14px 0;border-bottom:1px solid var(--border-default);}
.km-trip-head{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:9px;}
.km-trip-head-right{display:flex;align-items:center;gap:8px;}
.km-trip-km-pill{font-family:var(--font-mono);font-size:10.5px;font-weight:700;background:var(--coca-red);color:#FFF;padding:3px 9px;border:1.5px solid var(--border-strong);}
.km-trip-body{font-size:13px;line-height:1.5;color:var(--text-secondary);}
.km-trip-line{display:flex;gap:10px;align-items:flex-start;}
.km-trip-vline{margin-left:3.5px;height:11px;width:1px;background:var(--border-default);}
.km-minimap{position:relative;margin-top:10px;border:1.5px solid var(--border-strong);background:var(--bg-elevated);overflow:hidden;}
.km-minimap-label{position:absolute;top:6px;left:8px;right:8px;display:flex;justify-content:space-between;font-family:var(--font-mono);font-size:9px;color:var(--text-muted);text-transform:uppercase;}
.km-marquee{margin-top:28px;border-top:2px solid var(--border-strong);border-bottom:2px solid var(--border-strong);background:var(--bg-dark);color:#FFF;overflow:hidden;padding:8px 0;}
.km-marquee-track{display:flex;white-space:nowrap;animation:marquee 32s linear infinite;font-family:var(--font-mono);font-size:10.5px;text-transform:uppercase;}
@keyframes marquee{0%{transform:translateX(0);}100%{transform:translateX(-50%);}}
.km-footer{text-align:center;margin-top:22px;}
.km-footer p{font-size:9.5px;font-family:var(--font-mono);color:var(--text-muted);text-transform:uppercase;}
.km-coca{color:var(--coca-red);font-weight:700;}
.km-spin{animation:spin 1s linear infinite;}
@keyframes spin{from{transform:rotate(0deg);}to{transform:rotate(360deg);}}
.km-fade-in{animation:fadeIn 0.3s ease-out;}
@keyframes fadeIn{from{opacity:0;}to{opacity:1;}}
.km-fade-up{animation:fadeUp 0.35s var(--ease-ios) both;}
@keyframes fadeUp{from{opacity:0;transform:translateY(10px);}to{opacity:1;transform:translateY(0);}}
.km-alert{display:flex;align-items:center;gap:10px;padding:12px 14px;border:1.5px solid var(--border-strong);border-radius:var(--radius-md);margin-bottom:12px;font-size:13px;font-weight:500;}
.km-alert--err{border-color:var(--coca-red);background:rgba(230,26,39,0.04);color:var(--coca-red);}
.km-alert--ok{border-color:var(--status-green);background:rgba(22,163,74,0.04);color:var(--status-green);}
.km-odometer{background:var(--bg-dark);border:1.5px solid var(--border-strong);border-radius:var(--radius-lg);padding:16px 18px;margin:16px 0;display:flex;align-items:center;justify-content:space-between;position:relative;overflow:hidden;}
.km-odometer-num{display:flex;align-items:baseline;gap:6px;font-family:var(--font-mono);font-weight:700;font-size:36px;color:#FFF;}
.km-odometer-unit{font-size:14px;color:var(--coca-red);font-weight:800;}
.km-odometer-money{font-family:var(--font-mono);font-size:11px;color:var(--coca-red);font-weight:600;}
.km-splash{position:fixed;inset:0;background:var(--bg-dark);z-index:9999;display:flex;align-items:center;justify-content:center;transition:opacity 0.6s var(--ease-ios);}
.km-splash--exit{opacity:0;pointer-events:none;}
.km-splash-logo{font-family:var(--font-display);font-weight:900;font-size:48px;color:#FFF;letter-spacing:-0.05em;}
.km-splash-logo span:last-child{color:var(--coca-red);}
`;
