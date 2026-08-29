const cfg = window.UCFS_CONFIG || {};
const $ = id => document.getElementById(id);
let sb, me, study, currentCase;

function msg(el, text, ok=false){ el.innerHTML = text ? `<div class="${ok?'ok':'error'}">${text}</div>` : ''; }
function teamsLink(phone){ return `https://teams.microsoft.com/l/call/0/0?users=${encodeURIComponent('4:'+phone)}`; }
function qualtricsLink(c){
  const u = new URL(c.qualtrics_url || study.qualtrics_url);
  u.searchParams.set('UCFS_CASE_ID', c.external_case_id);
  u.searchParams.set('MODE','CATI');
  u.searchParams.set('INTERVIEWER_ID', me.id);
  return u.toString();
}

async function init(){
  if(!cfg.SUPABASE_URL || cfg.SUPABASE_URL.includes('YOUR_PROJECT')){
    msg($('loginMsg'),'Edit config.js with your Supabase URL and public anon/publishable key first.');
    return;
  }
  sb = supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY);
  const {data:{session}} = await sb.auth.getSession();
  if(session) await enter(session.user);
}

$('loginBtn').onclick = async()=>{
  msg($('loginMsg'),'');
  const emailValue = $('loginEmail').value.trim();
  const passwordValue = $('loginPassword').value;
  if (!emailValue || !passwordValue) {
    return msg($('loginMsg'),'Enter both email and password.');
  }
  const {data,error}=await sb.auth.signInWithPassword({
    email: emailValue,
    password: passwordValue
  });
  if(error) return msg($('loginMsg'),error.message);
  await enter(data.user);
};

$('logout').onclick=async()=>{await sb.auth.signOut();location.reload();};

async function enter(user){
  const {data:p,error:pe}=await sb.from('profiles').select('*').eq('id',user.id).single();
  if(pe) return msg($('loginMsg'),pe.message);
  me=p;
  const {data:s,error:se}=await sb.from('studies').select('*').eq('code',cfg.STUDY_CODE).single();
  if(se) return msg($('loginMsg'),se.message);
  study=s;
  $('loginCard').classList.add('hidden'); $('userBar').classList.remove('hidden');
  $('who').textContent=me.display_name; $('role').textContent=me.role; $('studyName').textContent=study.name;
  $('interviewerView').classList.remove('hidden');
  if(me.role!=='INTERVIEWER'){ $('managementView').classList.remove('hidden'); await loadManagement(); }
  await nextCase();
}

async function nextCase(){
  const {data,error}=await sb.rpc('reserve_next_case',{p_study:study.id});
  if(error){ msg($('caseMsg'),error.message); return; }
  currentCase = Array.isArray(data)?data[0]:data;
  if(!currentCase){
    $('caseBox').classList.add('hidden'); $('outcomeCard').classList.add('hidden');
    $('caseEmpty').classList.remove('hidden'); $('caseEmpty').textContent='No eligible case is currently available.';
    return;
  }
  $('caseEmpty').classList.add('hidden'); $('caseBox').classList.remove('hidden'); $('outcomeCard').classList.remove('hidden');
  $('caseId').textContent=currentCase.external_case_id; $('caseName').textContent=currentCase.display_name||'';
  $('casePhone').textContent=currentCase.phone; $('caseLanguage').textContent=currentCase.language||'';
  $('caseAttempts').textContent=currentCase.attempt_count; $('caseGender').textContent=currentCase.gender||'';
  $('caseStratum').textContent=currentCase.stratum||''; $('caseAge').textContent=currentCase.age||'';
  $('caseAgeBand').textContent=currentCase.age_band||''; $('caseCallback').textContent=currentCase.callback_at?new Date(currentCase.callback_at).toLocaleString():'—';
  $('teamsBtn').href=teamsLink(currentCase.phone); $('qualtricsBtn').href=qualtricsLink(currentCase);
  $('outcome').value=''; $('callbackAt').value=''; $('notes').value=''; msg($('caseMsg'),'');
}

$('releaseBtn').onclick=async()=>{
  if(!currentCase)return;
  const {error}=await sb.rpc('release_case',{p_case:currentCase.id});
  if(error)return msg($('caseMsg'),error.message);
  currentCase=null; await nextCase(); if(me.role!=='INTERVIEWER')await loadManagement();
};

$('saveNextBtn').onclick=async()=>{
  const out=$('outcome').value; if(!out)return msg($('caseMsg'),'Choose an outcome.');
  const callback = $('callbackAt').value ? new Date($('callbackAt').value).toISOString() : null;
  if(out==='CALLBACK' && !callback)return msg($('caseMsg'),'Enter a callback date and time.');
  const {error}=await sb.rpc('save_outcome',{p_case:currentCase.id,p_outcome:out,p_callback_at:callback,p_notes:$('notes').value||null});
  if(error)return msg($('caseMsg'),error.message);
  currentCase=null; await nextCase(); if(me.role!=='INTERVIEWER')await loadManagement();
};

async function loadManagement(){
  const {data:cases}=await sb.from('cases').select('id,status,reserved_by,callback_at,external_case_id,display_name');
  const counts={}; (cases||[]).forEach(c=>counts[c.status]=(counts[c.status]||0)+1);
  $('stats').innerHTML=['AVAILABLE','RESERVED','CALLBACK','COMPLETE','CLOSED'].map(k=>`<div class="card"><b>${k}</b><div style="font-size:28px">${counts[k]||0}</div></div>`).join('');
  $('qGender').checked=study.quota_gender_enabled; $('qStratum').checked=study.quota_stratum_enabled; $('qAge').checked=study.quota_age_enabled;

  const {data:profiles}=await sb.from('profiles').select('id,display_name');
  const names=Object.fromEntries((profiles||[]).map(p=>[p.id,p.display_name]));
  $('reservationRows').innerHTML=(cases||[]).filter(c=>c.reserved_by).map(c=>`<tr><td>${c.external_case_id}</td><td>${c.display_name||''}</td><td>${names[c.reserved_by]||c.reserved_by}</td><td>${c.status}</td><td>${c.callback_at?new Date(c.callback_at).toLocaleString():''}</td></tr>`).join('');

  const {data:q}=await sb.from('quota_targets').select('*').eq('study_id',study.id).order('dimension').order('category');
  renderQuotaTables(q||[]);
}

$('saveQuotaToggles').onclick=async()=>{
  const patch={quota_gender_enabled:$('qGender').checked,quota_stratum_enabled:$('qStratum').checked,quota_age_enabled:$('qAge').checked};
  const {data,error}=await sb.from('studies').update(patch).eq('id',study.id).select().single();
  if(error)return alert(error.message); study=data; alert('Quota toggles saved.');
};

function renderQuotaTables(q){
  const dims=[
    ['gender','Gender',['Female','Male']],
    ['stratum','Strata',["Nicosia Urban","Nicosia Rural","Limassol Urban","Limassol Rural","Larnaca Urban","Larnaca Rural","Paphos Urban","Paphos Rural","Famagusta Rural"]],
    ['age_band','Age bands',['18-22','23-27','28-32','33-37','38-42','43-47','48-52','53-57','58-62','63-67','68-72','73-77','78-82','83-87','88+']]
  ];
  const map=new Map(q.map(x=>[x.dimension+'|'+x.category,x]));
  $('quotaTables').innerHTML=dims.map(([d,title,cats])=>`
    <h4>${title}</h4><table><thead><tr><th>Category</th><th>Target</th><th></th></tr></thead><tbody>
    ${cats.map(cat=>{const x=map.get(d+'|'+cat);return `<tr><td>${cat}</td><td><input style="width:100px" type="number" min="0" id="qt_${btoa(unescape(encodeURIComponent(d+'|'+cat))).replaceAll('=','')}" value="${x?x.target:0}"></td><td><button class="secondary" onclick="saveQuota('${d}',${JSON.stringify(cat)})">Save</button></td></tr>`}).join('')}
    </tbody></table>`).join('');
}

window.saveQuota=async function(d,cat){
  const key=btoa(unescape(encodeURIComponent(d+'|'+cat))).replaceAll('=','');
  const target=parseInt($('qt_'+key).value||'0',10);
  const {error}=await sb.from('quota_targets').upsert({study_id:study.id,dimension:d,category:cat,target},{onConflict:'study_id,dimension,category'});
  if(error)alert(error.message); else alert('Quota target saved.');
};

init();
