import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, ReferenceLine,
} from "recharts";

const COLORS = [
  "#f59e0b","#10b981","#3b82f6","#ef4444","#8b5cf6",
  "#ec4899","#06b6d4","#84cc16","#f97316","#6366f1",
  "#14b8a6","#f43f5e","#a78bfa","#34d399","#fbbf24",
  "#60a5fa","#fb7185","#4ade80","#c084fc","#38bdf8",
];

const DARK = {
  
  bg:"#0a0a0f", surface:"#0d0d15", surface2:"#0f0f18",
  border:"#1e2030", border2:"#13131a",
  text:"#e8e4d9", textMuted:"#9ca3af", textDim:"#6b7280",
  textFaint:"#374151", textFaint2:"#1f2937",
  gridLine:"#1a1a24", inputBg:"#0d0d15",
  headerBg:"linear-gradient(90deg,#0a0a0f,#0d1117)",
  sunColor:"#fbbf24", dotBg:"#2a2d3e",
};
const LIGHT = {
  bg:"#f5f5f0", surface:"#ffffff", surface2:"#f0f0eb",
  border:"#d1d5db", border2:"#e5e7eb",
  text:"#1a1a2e", textMuted:"#4b5563", textDim:"#6b7280",
  textFaint:"#9ca3af", textFaint2:"#d1d5db",
  gridLine:"#e5e7eb", inputBg:"#ffffff",
  headerBg:"linear-gradient(90deg,#f5f5f0,#eeeee8)",
  sunColor:"#d97706", dotBg:"#d1d5db",
};

const toColLabel = (i) => {
  let label = "", n = i;
  do { label = String.fromCharCode(65 + (n % 26)) + label; n = Math.floor(n / 26) - 1; } while (n >= 0);
  return label;
};

function computeJVParams(voltages, currents) {
  let isc = null;
  for (let i = 0; i < voltages.length - 1; i++) {
    const v1=voltages[i],v2=voltages[i+1],j1=currents[i],j2=currents[i+1];
    if ((v1<=0&&v2>=0)||(v1>=0&&v2<=0)){isc=j1+((0-v1)/(v2-v1))*(j2-j1);break;}
  }
  if (isc===null){let mi=0;for(let i=1;i<voltages.length;i++)if(Math.abs(voltages[i])<Math.abs(voltages[mi]))mi=i;isc=currents[mi];}
  let voc=null;
  for(let i=0;i<voltages.length-1;i++){
    const j1=currents[i],j2=currents[i+1],v1=voltages[i],v2=voltages[i+1];
    if((j1<=0&&j2>=0)||(j1>=0&&j2<=0)){voc=v1+((0-j1)/(j2-j1))*(v2-v1);break;}
  }
  if(voc===null){let mi=0;for(let i=1;i<currents.length;i++)if(Math.abs(currents[i])<Math.abs(currents[mi]))mi=i;voc=voltages[mi];}
  let pmax=0;
  for(let i=0;i<voltages.length;i++)
    if(voltages[i]>0&&currents[i]<0)pmax=Math.max(pmax,Math.abs(voltages[i]*currents[i]));
  const ff=(voc&&isc)?pmax/(Math.abs(voc)*Math.abs(isc)):0;
  const pce=(Math.abs(isc)*Math.abs(voc)*ff)/100;
  return {isc:Math.abs(isc),voc:Math.abs(voc),ff,pce,pmax};
}

function parseFileToSample(file, vCol, jCol, diameter, colorIndex) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, {type:"array"});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, {header:1});
      if (!json.length){resolve(null);return;}
      const numCols = Math.max(...json.map(r=>r.length));
      const columns = Array.from({length:numCols},(_,i)=>toColLabel(i));
      const vIdx=columns.indexOf(vCol), jIdx=columns.indexOf(jCol);
      if(vIdx===-1||jIdx===-1){resolve(null);return;}
      const diam=parseFloat(diameter);
      const areaCm2=diam>0?Math.PI*Math.pow((diam/2)*0.1,2):1;
      const voltages=[],currents=[];
      for(const row of json){
        const v=parseFloat(row[vIdx]),iRaw=parseFloat(row[jIdx]);
        if(!isNaN(v)&&!isNaN(iRaw)){voltages.push(v);currents.push((iRaw*1000)/areaCm2);}
      }
      if(voltages.length<2){resolve(null);return;}
      const params=computeJVParams(voltages,currents);
      const chartData=voltages.map((v,i)=>({v,j:currents[i]}));
      const name=file.name.replace(/\.[^/.]+$/,"");
      resolve({id:`${Date.now()}-${Math.random()}`,name,params,chartData,color:COLORS[colorIndex%COLORS.length],visible:true});
    };
    reader.readAsArrayBuffer(file);
  });
}

function getSampleGroup(name) {
  const m = name.match(/^(sample\d+)/i);
  return m ? m[1].toLowerCase() : null;
}

function RightPanel({samples, toggleVisible, toggleAll, removeSample, T}) {
  const [collapsed, setCollapsed] = useState({});
  const groups={}, ungrouped=[];
  for(const s of samples){
    const g=getSampleGroup(s.name);
    if(g){if(!groups[g])groups[g]=[];groups[g].push(s);}
    else ungrouped.push(s);
  }
  const groupKeys=Object.keys(groups).sort((a,b)=>parseInt(a.replace(/\D/g,""))-parseInt(b.replace(/\D/g,"")));
  const btnStyle={background:T.surface2,border:`1px solid ${T.border}`,color:T.textMuted,borderRadius:5,padding:"4px 0",fontSize:10,cursor:"pointer",fontFamily:"inherit",flex:1};

  return (
    <div style={{width:230,flexShrink:0,borderLeft:`1px solid ${T.border}`,padding:"20px 14px",display:"flex",flexDirection:"column",gap:10,overflowY:"auto",background:T.bg}}>
      <div style={{fontSize:10,letterSpacing:"0.2em",color:T.textDim,marginBottom:2}}>SAMPLES ({samples.length})</div>
      <div style={{display:"flex",gap:6,marginBottom:4}}>
        <button onClick={()=>toggleAll(true)} style={btnStyle}>All on</button>
        <button onClick={()=>toggleAll(false)} style={btnStyle}>All off</button>
      </div>
      {groupKeys.map(gKey=>{
        const gs=groups[gKey];
        const allOn=gs.every(s=>s.visible),allOff=gs.every(s=>!s.visible);
        const isCollapsed=collapsed[gKey];
        return (
          <div key={gKey}>
            <div style={{display:"flex",alignItems:"center",gap:6,padding:"6px 8px",background:T.surface2,border:`1px solid ${T.border}`,borderRadius:isCollapsed?6:"6px 6px 0 0",cursor:"pointer"}}>
              <span onClick={()=>setCollapsed(c=>({...c,[gKey]:!c[gKey]}))} style={{fontSize:9,color:T.textDim,userSelect:"none",width:10}}>{isCollapsed?"▶":"▼"}</span>
              <span onClick={()=>setCollapsed(c=>({...c,[gKey]:!c[gKey]}))} style={{fontSize:11,color:T.text,fontWeight:700,flex:1,letterSpacing:"0.05em"}}>{gKey}</span>
              <span onClick={()=>gs.forEach(s=>{if(s.visible!==!allOn)toggleVisible(s.id);})}
                style={{fontSize:9,color:allOn?"#f59e0b":T.textFaint,cursor:"pointer",userSelect:"none",border:`1px solid ${allOn?"#f59e0b44":T.border}`,borderRadius:3,padding:"2px 5px"}}>
                {allOff?"off":allOn?"on":"—"}
              </span>
            </div>
            {!isCollapsed&&(
              <div style={{border:`1px solid ${T.border}`,borderTop:"none",borderRadius:"0 0 6px 6px",overflow:"hidden"}}>
                {gs.map((s,idx)=>(
                  <div key={s.id} onClick={()=>toggleVisible(s.id)}
                    style={{display:"flex",alignItems:"center",gap:7,padding:"6px 8px 6px 18px",background:s.visible?T.surface:"transparent",borderTop:idx>0?`1px solid ${T.border2}`:"none",cursor:"pointer"}}>
                    <div style={{width:8,height:8,borderRadius:"50%",flexShrink:0,background:s.visible?s.color:T.dotBg,boxShadow:s.visible?`0 0 5px ${s.color}88`:"none"}}/>
                    <span style={{fontSize:10,flex:1,color:s.visible?T.text:T.textDim,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={s.name}>{s.name}</span>
                    <button onClick={e=>{e.stopPropagation();removeSample(s.id);}} style={{background:"none",border:"none",color:T.textFaint,cursor:"pointer",fontSize:11,padding:0,fontFamily:"inherit",flexShrink:0}}>✕</button>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })}
      {ungrouped.length>0&&(
        <div style={{display:"flex",flexDirection:"column",gap:5}}>
          {groupKeys.length>0&&<div style={{fontSize:10,color:T.textFaint,letterSpacing:"0.1em",marginTop:4}}>OTHER</div>}
          {ungrouped.map(s=>(
            <div key={s.id} onClick={()=>toggleVisible(s.id)}
              style={{display:"flex",alignItems:"center",gap:7,background:s.visible?T.surface:"transparent",border:`1px solid ${s.visible?s.color+"44":T.border2}`,borderRadius:6,padding:"7px 10px",cursor:"pointer"}}>
              <div style={{width:8,height:8,borderRadius:"50%",flexShrink:0,background:s.visible?s.color:T.dotBg,boxShadow:s.visible?`0 0 5px ${s.color}88`:"none"}}/>
              <span style={{fontSize:11,color:s.visible?T.text:T.textDim,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={s.name}>{s.name}</span>
              <button onClick={e=>{e.stopPropagation();removeSample(s.id);}} style={{background:"none",border:"none",color:T.textFaint,cursor:"pointer",fontSize:12,padding:0,fontFamily:"inherit",flexShrink:0}}>✕</button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

const CHART_MARGIN={top:8,right:24,left:40,bottom:20};

export default function JVAnalyzer() {
  const [dark, setDark] = useState(true);
  const T = dark ? DARK : LIGHT;

  const [samples, setSamples] = useState([]);
  const [vCol, setVCol] = useState("A");
  const [jCol, setJCol] = useState("B");
  const [diameter, setDiameter] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [loading, setLoading] = useState(false);
  const [hoveredId, setHoveredId] = useState(null);
  const fileRef = useRef();

  const handleFiles = async (files) => {
    setLoading(true);
    const fileArray = Array.from(files).filter(f=>f.name.match(/\.(csv|xlsx|xls)$/i));
    const newSamples = [];
    for(let i=0;i<fileArray.length;i++){
      const s = await parseFileToSample(fileArray[i],vCol,jCol,diameter,samples.length+i);
      if(s) newSamples.push(s);
    }
    setSamples(prev=>[...prev,...newSamples]);
    setLoading(false);
  };

  const handleFileInput = e => handleFiles(e.target.files);
  const handleDrop = e => {e.preventDefault();setDragOver(false);handleFiles(e.dataTransfer.files);};
  const toggleVisible = id => setSamples(s=>s.map(x=>x.id===id?{...x,visible:!x.visible}:x));
  const toggleAll = vis => setSamples(s=>s.map(x=>({...x,visible:vis})));
  const removeSample = id => setSamples(s=>s.filter(x=>x.id!==id));
  const clearAll = () => setSamples([]);
  const visibleSamples = samples.filter(s=>s.visible);

  const areaCm2Display = (()=>{const d=parseFloat(diameter);return d>0?(Math.PI*Math.pow((d/2)*0.1,2)).toFixed(4):null;})();

  return (
    <div style={{minHeight:"100vh",background:T.bg,color:T.text,fontFamily:"'DM Mono','Fira Code',monospace",display:"flex",flexDirection:"column",transition:"background 0.2s,color 0.2s"}}>
      {/* Header */}
      <div style={{borderBottom:`1px solid ${T.border}`,padding:"20px 32px",display:"flex",alignItems:"center",gap:14,background:T.headerBg}}>
        <div style={{width:32,height:32,borderRadius:"50%",flexShrink:0,background:`radial-gradient(circle at 40% 40%,${T.sunColor},${T.sunColor}88)`,boxShadow:`0 0 18px ${T.sunColor}66`}}/>
        <div>
          <div style={{fontSize:18,fontWeight:700,letterSpacing:"0.05em",color:T.text}}>JV CURVE ANALYZER</div>
          <div style={{fontSize:10,color:T.textDim,letterSpacing:"0.15em"}}>SOLAR CELL · AM1.5G · BATCH MODE</div>
        </div>
        <div style={{marginLeft:"auto",display:"flex",gap:10,alignItems:"center"}}>
          <button onClick={()=>setDark(d=>!d)} style={{background:T.surface2,border:`1px solid ${T.border}`,color:T.textMuted,borderRadius:20,padding:"5px 14px",fontSize:13,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:6,transition:"all 0.2s"}}>
            {dark?"☀️ Light":"🌙 Dark"}
          </button>
          {samples.length>0&&<>
            <span style={{fontSize:11,color:T.textDim}}>{samples.length} sample{samples.length>1?"s":""} loaded</span>
            <button onClick={clearAll} style={{background:T.surface2,border:`1px solid ${T.border}`,color:T.textMuted,borderRadius:6,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>Clear all</button>
          </>}
        </div>
      </div>

      <div style={{display:"flex",flex:1}}>
        {/* Left Panel */}
        <div style={{width:280,flexShrink:0,borderRight:`1px solid ${T.border}`,padding:"24px 20px",display:"flex",flexDirection:"column",gap:22,background:T.bg}}>
          {/* Upload */}
          <div>
            <div style={{fontSize:10,letterSpacing:"0.2em",color:T.textDim,marginBottom:10}}>01 · UPLOAD FILES</div>
            <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)} onDrop={handleDrop} onClick={()=>fileRef.current.click()}
              style={{border:`1px dashed ${dragOver?"#f59e0b":T.border}`,borderRadius:8,padding:"22px 12px",textAlign:"center",cursor:"pointer",background:dragOver?"#f59e0b08":T.surface,transition:"all 0.2s"}}>
              {loading?<div style={{fontSize:12,color:"#f59e0b"}}>⏳ Processing...</div>:<>
                <div style={{fontSize:24,marginBottom:6}}>☀️</div>
                <div style={{fontSize:12,color:T.textMuted}}>Drop multiple CSV/Excel files<br/><span style={{color:T.textDim,fontSize:11}}>or click to browse</span></div>
              </>}
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" multiple onChange={handleFileInput} style={{display:"none"}}/>
            </div>
          </div>

          {/* Column mapping */}
          <div>
            <div style={{fontSize:10,letterSpacing:"0.2em",color:T.textDim,marginBottom:10}}>02 · COLUMN MAPPING</div>
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              <label style={{fontSize:11,color:T.textDim}}>VOLTAGE COLUMN
                <input value={vCol} onChange={e=>setVCol(e.target.value.toUpperCase())}
                  style={{display:"block",width:"100%",marginTop:4,boxSizing:"border-box",background:T.inputBg,border:`1px solid ${T.border}`,color:"#f59e0b",borderRadius:6,padding:"7px 10px",fontSize:14,fontFamily:"inherit",fontWeight:700}}/>
              </label>
              <label style={{fontSize:11,color:T.textDim}}>CURRENT COLUMN (A)
                <input value={jCol} onChange={e=>setJCol(e.target.value.toUpperCase())}
                  style={{display:"block",width:"100%",marginTop:4,boxSizing:"border-box",background:T.inputBg,border:`1px solid ${T.border}`,color:"#3b82f6",borderRadius:6,padding:"7px 10px",fontSize:14,fontFamily:"inherit",fontWeight:700}}/>
              </label>
            </div>
          </div>

          {/* Diameter */}
          <div>
            <div style={{fontSize:10,letterSpacing:"0.2em",color:T.textDim,marginBottom:10}}>03 · CELL DIAMETER</div>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <input type="number" min="0" step="0.01" value={diameter} onChange={e=>setDiameter(e.target.value)} placeholder="e.g. 6"
                style={{flex:1,background:T.inputBg,border:`1px solid ${T.border}`,color:T.text,borderRadius:6,padding:"7px 10px",fontSize:13,fontFamily:"inherit"}}/>
              <span style={{fontSize:12,color:T.textDim}}>mm</span>
            </div>
            {areaCm2Display&&<div style={{fontSize:10,color:"#f59e0b",marginTop:6}}>Area = {areaCm2Display} cm²</div>}
            {!diameter&&<div style={{fontSize:10,color:T.textFaint,marginTop:6}}>Leave blank if already mA/cm²</div>}
          </div>
        </div>

        {/* Center */}
        <div style={{flex:1,padding:"24px 28px",display:"flex",flexDirection:"column",gap:24,minWidth:0}}>
          {samples.length===0?(
            <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:12,color:T.textFaint}}>
              <div style={{fontSize:56}}>📂</div>
              <div style={{fontSize:13,letterSpacing:"0.08em",textAlign:"center"}}>
                Drop your CSV files to get started<br/>
                <span style={{fontSize:11,color:T.textFaint2}}>Multiple files supported</span>
              </div>
            </div>
          ):(<>
            {/* Chart */}
            <div>
              <div style={{fontSize:10,letterSpacing:"0.2em",color:T.textDim,marginBottom:12}}>
                JV CURVES {visibleSamples.length<samples.length&&`(${visibleSamples.length}/${samples.length} shown)`}
              </div>
              <div style={{background:T.surface,border:`1px solid ${T.border}`,borderRadius:10,padding:"20px 8px 12px 8px"}}>
                {visibleSamples.length===0
                  ?<div style={{height:280,display:"flex",alignItems:"center",justifyContent:"center",color:T.textFaint,fontSize:12}}>No samples selected</div>
                  :<ResponsiveContainer width="100%" height={280}>
                    <LineChart margin={CHART_MARGIN} onMouseLeave={()=>setHoveredId(null)}>
                      <CartesianGrid strokeDasharray="3 3" stroke={T.gridLine}/>
                      <XAxis dataKey="v" type="number" domain={["auto","auto"]}
                        tick={{fill:T.textDim,fontSize:10}}
                        label={{value:"Voltage (V)",position:"insideBottom",offset:-10,fill:T.textDim,fontSize:11}}/>
                      <YAxis
                        tickCount={6}
                        tickFormatter={v=>Number(v).toFixed(2)}
                        tick={{fill:T.textDim,fontSize:10}}
                        label={{value:"J (mA/cm2)",angle:-90,position:"insideLeft",offset:10,fill:T.textDim,fontSize:11}}/>
                      <Tooltip
                        cursor={{stroke:T.border,strokeWidth:1,strokeDasharray:"4 4"}}
                        content={({active,payload,label})=>{
                          if(!active||!hoveredId||!payload?.length) return null;
                          const entry=payload.find(p=>{
                            const s=visibleSamples.find(s=>s.name===p.name);
                            return s&&s.id===hoveredId;
                          });
                          if(!entry) return null;
                          return (
                            <div style={{background:T.surface,border:`1px solid ${entry.color}88`,borderRadius:6,padding:"7px 11px",fontSize:11,boxShadow:`0 0 12px ${entry.color}22`,pointerEvents:"none",whiteSpace:"nowrap"}}>
                              <div style={{color:T.textDim,fontSize:10,marginBottom:3}}>V = <span style={{color:T.text,fontWeight:700}}>{Number(label).toFixed(4)}</span> V</div>
                              <div style={{color:entry.color,fontWeight:600,marginBottom:2}}>{entry.name}</div>
                              <div style={{color:entry.color}}>J = {Number(entry.value).toFixed(4)} mA/cm²</div>
                            </div>
                          );
                        }}
                      />
                      <ReferenceLine y={0} stroke={T.border}/>
                      <ReferenceLine x={0} stroke={T.border}/>
                      {visibleSamples.map(s=>(
                        <Line key={s.id} data={s.chartData} dataKey="j" name={s.name}
                          stroke={s.color} dot={false} strokeWidth={2}
                          type="monotone" isAnimationActive={false}
                          activeDot={hoveredId===s.id?{r:5,fill:s.color,stroke:T.surface,strokeWidth:2}:false}
                          onMouseEnter={()=>setHoveredId(s.id)}/>
                      ))}
                    </LineChart>
                  </ResponsiveContainer>
                }
              </div>
            </div>

            {/* Table */}
            <div>
              <div style={{fontSize:10,letterSpacing:"0.2em",color:T.textDim,marginBottom:12}}>PARAMETERS — ALL SAMPLES</div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead>
                    <tr style={{borderBottom:`1px solid ${T.border}`}}>
                      {["","File","V_OC (V)","J_SC (mA/cm²)","FF (%)","PCE (%)"].map((h,i)=>(
                        <th key={i} style={{padding:"7px 12px",textAlign:i>1?"right":"left",fontSize:10,letterSpacing:"0.12em",color:T.textDim,fontWeight:600}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {samples.map(s=>(
                      <tr key={s.id} style={{borderBottom:`1px solid ${T.border2}`,opacity:s.visible?1:0.4}}>
                        <td style={{padding:"8px 12px"}}><div style={{width:8,height:8,borderRadius:"50%",background:s.color}}/></td>
                        <td style={{padding:"8px 12px",color:T.textMuted,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.name}</td>
                        <td style={{padding:"8px 12px",textAlign:"right",color:"#10b981",fontWeight:600}}>{s.params.voc.toFixed(4)}</td>
                        <td style={{padding:"8px 12px",textAlign:"right",color:"#3b82f6",fontWeight:600}}>{s.params.isc.toFixed(4)}</td>
                        <td style={{padding:"8px 12px",textAlign:"right",color:"#f59e0b",fontWeight:600}}>{(s.params.ff*100).toFixed(2)}%</td>
                        <td style={{padding:"8px 12px",textAlign:"right",color:"#ef4444",fontWeight:700}}>{(s.params.pce*100).toFixed(2)}%</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{fontSize:10,color:T.textFaint,marginTop:6}}>* PCE assumes Pin = 100 mW/cm² (AM1.5G)</div>
            </div>
          </>)}
        </div>

        {/* Right Panel */}
        {samples.length>0&&<RightPanel samples={samples} toggleVisible={toggleVisible} toggleAll={toggleAll} removeSample={removeSample} T={T}/>}
      </div>
    </div>
  );
}
