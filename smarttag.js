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
  console.log("[LOG UI]", t);
  const m=$("#log");
  if(!m){ console.warn("[NO LOG DIV]"); return; }
  m.className="msg"+(e?" err":"");
  m.innerHTML=t;
};

const sb=o=>$("#bar").style.opacity=o?"1":"0";
const st=o=>$("#typing").classList.toggle("hidden",!o);

// ✅ DEBUG INIT
Office.onReady(()=>{ 
  console.log("✅ Office ready, init() running...");
  init();
});

function init(){
  console.log("🔧 Init UI listeners");
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
    console.log("🔁 Refresh clicked");
    tags=null;
    sugs=[];
    hide();
    log("🔄 Đang làm mới dữ liệu…");
    await fetchTags();
  });

  fetchTags();
}

async function fetchTags(){
  if(tags){
    console.log("📦 Tags cache hit", tags.length);
    return tags;
  }

  sb(1); st(1);
  log("⏳ Đang tải dữ liệu từ SharePoint…");
  console.log("🌐 Fetching SP list:", CONFIG.SITE_URL);

  try {
    const url=`${CONFIG.SITE_URL}/_api/web/lists/getbytitle('${CONFIG.LIST_NAME}')/items?$select=Title,Value,Desc&$top=${CONFIG.TOP}`;
    console.log("➡️ API:", url);

    const res=await fetch(url,{
      headers:{Accept:"application/json;odata=verbose"},
      credentials:"include"
    });

    console.log("⬅️ Response status:", res.status);

    if(!res.ok){
      let errText="";
      try { errText = await res.text(); } catch {}

      log(`❌ SharePoint lỗi<br>HTTP: ${res.status}<br><small>${errText}</small>`,1);
      throw new Error("SharePoint fetch failed");
    }

    const json = await res.json();
    tags = json?.d?.results || [];
    console.log(`✅ Loaded tags:`, tags);

    if(!tags.length){
      log(`⚠️ List <b>${CONFIG.LIST_NAME}</b> rỗng hoặc không có quyền`,1);
      return [];
    }

    log(`✅ Tải ${tags.length} tags thành công`);
    return tags;

  } catch(e) {
    console.error("❌ FetchTags ERROR:", e);
    log(`❌ Không lấy được dữ liệu SP<br><small>${e.message}</small>`,1);
  } finally {
    sb(0); st(0);
  }
}

const deb=(fn,ms)=>{let t; return (...a)=>{clearTimeout(t); t=setTimeout(()=>fn(...a),ms);}};

async function type(e){
  const v=e.target.value;
  console.log("⌨️ Input:", v);
  
  if(!v.startsWith("@")){
    console.log("⛔ Không phải @, hide suggest");
    return hide();
  }

  const k=v.slice(1).toLowerCase();
  console.log(`🔍 Detect @ — key = "${k}"`);

  const t = await fetchTags();
  sugs = !k ? t.slice(0,20) : t.filter(x => 
    `${x.Title} ${x.Value} ${x.Desc}`.toLowerCase().includes(k)
  ).slice(0,20);

  console.log("📋 Suggestions:", sugs);

  render(sugs);
}

function render(arr){
  console.log("🎨 Render suggestions:", arr.length);

  const b=$("#suggest");
  const input=$("#ipt");
  si=0;

  if(!arr.length){
    console.log("🚫 No suggestions, hiding box");
    return hide();
  }

  b.innerHTML = arr.map((t,i)=>{
    const title=esc(t.Title);
    const desc=esc(t.Desc||"Không có mô tả");
    const val=esc(t.Value||"");
    return `
      <div class="s-item ${i===0?'active':''}" role="option"
           id="s-option-${i}" aria-selected="${i===0}"
           data-i="${i}" data-value="${val}">
        <div class="s-item-title">${title}</div>
        <small>${desc}</small>
      </div>`;
  }).join("");

  b.classList.remove("hidden");
  input.setAttribute("aria-expanded","true");
  input.setAttribute("aria-activedescendant","s-option-0");

  [...b.children].forEach(el=>{
    el.onclick=()=>{
      console.log("🖱️ Click select", el.dataset.i);
      si=+el.dataset.i;
      select();
    };
  });
}

function hide(){
  console.log("🙈 Hide suggest");
  const s=$("#suggest");
  s.classList.add("hidden");
  s.innerHTML="";
  $("#ipt").setAttribute("aria-expanded","false");
}

function key(e){
  if($("#suggest").classList.contains("hidden")) return;
  console.log("⚡ Key:", e.key);

  if(e.key==="ArrowDown"){e.preventDefault(); move(1);}
  else if(e.key==="ArrowUp"){e.preventDefault(); move(-1);}
  else if(e.key==="Enter"){e.preventDefault(); select();}
}

function move(d){
  console.log("🔄 Move:", d);
  const it=[...document.querySelectorAll(".s-item")];
  si=(si+d+it.length)%it.length;
  it.forEach((el,i)=>{
    el.classList.toggle("active",i===si);
    el.setAttribute("aria-selected",i===si);
  });
}

async function select(){
  const t=sugs[si];
  console.log("✅ SELECT:", t);

  hide();
  $("#cards").innerHTML=`
    <article class="card">
      <div class="card-header">
        <span class="card-pill">SharePoint</span>
        <span class="card-code">${esc(t.Value)}</span>
      </div>
      <h3>${esc(t.Title)}</h3>
      <p>${esc(t.Desc||"Không có mô tả")}</p>
    </article>`;

  await insert(t.Value);
}

async function insert(v){
  console.log("✏️ Insert:", v);
  sb(1); st(1);

  try{
    await Word.run(async ctx=>{
      ctx.document.getSelection().insertText(v,Word.InsertLocation.replace);
      await ctx.sync();
    });
    log(`✍️ Đã chèn: <b>${esc(v)}</b>`);
  } catch(e) {
    console.warn("⚠️ Word.insert failed, fallback clipboard", e);
    await navigator.clipboard.writeText(v);
    log(`📋 Sao chép: <b>${esc(v)}</b>`);
  }

  sb(0); st(0);
}
