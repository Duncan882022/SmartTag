const CONFIG={
  SITE_URL:"https://1c0ggy.sharepoint.com/sites/qlnb",
  LIST_NAME:"TagLibrary",
  TOP:500,
  DELAY:120
};

let tags=null,si=0,sugs=[];
const $=s=>document.querySelector(s);
const esc=s=>String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"})[m]);
const log=(t,e=false)=>{
  const m=$("#log");
  m.className="msg"+(e?" err":"");
  m.innerHTML=t;
};

const sb=o=>$("#bar").style.opacity=o?"1":"0";
const st=o=>$("#typing").classList.toggle("hidden",!o);

Office.onReady(()=>init());

function init(){
  const i=$("#ipt");
  i.setAttribute("aria-autocomplete","list");
  i.setAttribute("aria-haspopup","listbox");
  i.setAttribute("aria-controls","suggest");
  i.setAttribute("aria-expanded","false");
  i.addEventListener("input",deb(type,CONFIG.DELAY));
  i.addEventListener("keydown",key);
  i.focus();
  const r=$("#refresh");
  if(r) r.addEventListener("click",async ()=>{
    tags=null;
    sugs=[];
    hide();
    log("🔄 Đang làm mới dữ liệu…");
    await fetchTags();
  });
  fetchTags();
}

async function fetchTags(){
  if(tags) return tags;

  sb(1); st(1);
  log("⏳ Đang tải dữ liệu từ SharePoint…");

  try {
    const url=`${CONFIG.SITE_URL}/_api/web/lists/getbytitle('${CONFIG.LIST_NAME}')/items?$select=Title,Value,Desc&$top=${CONFIG.TOP}`;
    const res=await fetch(url,{
      headers:{Accept:"application/json;odata=verbose"},
      credentials:"include"
    });

    if(!res.ok){
      let errText="";
      try { errText = await res.text(); } catch {}

      log(
        `❌ Lỗi SharePoint<br>
         HTTP: ${res.status}<br>
         URL: ${CONFIG.SITE_URL}<br><br>
         <small>${errText}</small>`,1
      );
      throw new Error("SharePoint fetch failed");
    }

    const json = await res.json();
    tags = json?.d?.results || [];

    if(!tags.length){
      log(`⚠️ Danh sách <b>${CONFIG.LIST_NAME}</b> rỗng hoặc không có quyền`,1);
      return [];
    }

    log(`✅ Tải ${tags.length} tags thành công`);
    return tags;

  } catch(e) {
    log(`❌ Không lấy được dữ liệu SP<br><small>${e.message}</small>`,1);
  } finally {
    sb(0); st(0);
  }
}

const deb=(fn,ms)=>{let t; return (...a)=>{clearTimeout(t); t=setTimeout(()=>fn(...a),ms);}};

async function type(e){
  const v=e.target.value;
  if(!v.startsWith("@")) return hide();

  const k=v.slice(1).toLowerCase();
  const t = await fetchTags();
  sugs = !k ? t.slice(0,20) : t.filter(x => `${x.Title} ${x.Value} ${x.Desc}`.toLowerCase().includes(k)).slice(0,20);
  render(sugs);
}

function render(arr){
  const b=$("#suggest");
  const input=$("#ipt");
  si=0;
  if(!arr.length) return hide();

  b.innerHTML = arr.map((t,i)=>{
    const title=esc(t.Title);
    const desc=esc(t.Desc||"Không có mô tả");
    const val=esc(t.Value||"");
    return `
      <div class="s-item ${i===0?'active':''}" id="s-option-${i}" role="option" aria-selected="${i===0}" data-i="${i}" data-value="${val}">
        <div class="s-item-title">${title}</div>
        <small>${desc}</small>
      </div>`;
  }).join("");

  b.classList.remove("hidden");
  b.setAttribute("aria-hidden","false");
  input.setAttribute("aria-expanded","true");
  input.setAttribute("aria-activedescendant","s-option-0");
  b.scrollTop=0;
  [...b.children].forEach(el=>{
    el.onclick=()=>{si=+el.dataset.i;select();};
  });
}

function hide(){
  const s=$("#suggest");
  const input=$("#ipt");
  s.classList.add("hidden");
  s.setAttribute("aria-hidden","true");
  s.innerHTML="";
  input.setAttribute("aria-expanded","false");
  input.removeAttribute("aria-activedescendant");
}

function key(e){
  if($("#suggest").classList.contains("hidden")) return;
  if(e.key==="ArrowDown"){e.preventDefault();move(1);}
  else if(e.key==="ArrowUp"){e.preventDefault();move(-1);}
  else if(e.key==="Enter"){e.preventDefault();select();}
}

function move(d){
  const it=[...document.querySelectorAll(".s-item")];
  si=(si+d+it.length)%it.length;
  const input=$("#ipt");
  it.forEach((el,i)=>{
    const active=i===si;
    el.classList.toggle("active",active);
    el.setAttribute("aria-selected",active);
    if(active) input.setAttribute("aria-activedescendant",el.id);
  });
}

async function select(){
  const t=sugs[si];
  hide();
  const title=esc(t.Title);
  const desc=esc(t.Desc||"Không có mô tả");
  const value=t.Value;
  const displayVal=esc(value);
  $("#cards").innerHTML=`
    <article class="card">
      <div class="card-header">
        <span class="card-pill">SharePoint</span>
        <span class="card-code">${displayVal}</span>
      </div>
      <h3>${title}</h3>
      <p>${desc}</p>
    </article>`;
  await insert(value);
}

async function insert(v){
  sb(1); st(1);
  try{
    await Word.run(async ctx=>{
      ctx.document.getSelection().insertText(v,Word.InsertLocation.replace);
      await ctx.sync();
    });
    log(`✍️ Đã chèn: <b>${esc(v)}</b>`);
  } catch {
    await navigator.clipboard.writeText(v);
    log(`📋 Sao chép: <b>${esc(v)}</b>`);
  }
  sb(0); st(0);
}
