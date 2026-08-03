
let poles = [];
let currentIndex = null;

// Paste your deployed Apps Script Web App URL here:
const API_URL = "https://script.google.com/macros/s/AKfycbyQgr7hWXYSj7D6yXWTm1yhm1UMest-c910MFo0OoesvW-ceiJMdRXa6ZyXxiroDAcV/exec";

const $ = s => document.querySelector(s);

async function api(action, payload={}){
  if(!API_URL || API_URL.includes("PASTE_")) return null;
  const res = await fetch(API_URL, {
    method: "POST",
    headers: {"Content-Type":"text/plain;charset=utf-8"},
    body: JSON.stringify({action, ...payload})
  });
  if(!res.ok) throw new Error("API error");
  return await res.json();
}

function loadLocalState(){
  const saved = JSON.parse(localStorage.getItem('makeReadyState') || '{}');
  poles = poles.map(p => ({...p, ...(saved[p.projectPole] || {})}));
}
function saveLocalState(){
  const obj = {};
  poles.forEach(p => obj[p.projectPole] = {
    status:p.status, actualHoa:p.actualHoa, heightChanged:p.heightChanged,
    measuredDistance:p.measuredDistance, trimmingNotes:p.trimmingNotes,
    fieldNotes:p.fieldNotes, beforePhotoCount:p.beforePhotoCount||0,
    afterPhotoCount:p.afterPhotoCount||0
  });
  localStorage.setItem('makeReadyState', JSON.stringify(obj));
}
function badgeClass(s){return s==='Completed'?'completed':s==='In progress'?'progress':s==='Problem'?'problem':'not'}
function render(){
  const q = $('#search').value.toLowerCase();
  const sf = $('#statusFilter').value;
  const filtered = poles.filter(p => {
    const hit = (`${p.projectPole} ${p.poleId}`).toLowerCase().includes(q);
    return hit && (!sf || p.status===sf);
  });
  $('#list').innerHTML = filtered.map(p => `
    <article class="pole" onclick="openPole(${p.projectPole})">
      <div class="pole-top">
        <div><h3>Pole ${p.projectPole}</h3><div class="muted">${p.poleId || 'No pole ID'}</div></div>
        <span class="badge ${badgeClass(p.status)}">${p.status}</span>
      </div>
      <div class="desc">${p.description}</div>
    </article>`).join('');
  const c = s => poles.filter(p=>p.status===s).length;
  $('#stats').innerHTML = `
    <div class="stat"><b>${poles.length}</b>Всего</div>
    <div class="stat"><b>${c('Not started')}</b>Не начато</div>
    <div class="stat"><b>${c('In progress')}</b>В работе</div>
    <div class="stat"><b>${c('Completed')}</b>Готово</div>`;
}

window.openPole = function(projectPole){
  currentIndex = poles.findIndex(p=>p.projectPole===projectPole);
  const p = poles[currentIndex];
  $('#dlgTitle').textContent = `Pole ${p.projectPole}`;
  $('#dlgId').textContent = p.poleId || 'No pole ID';
  $('#navigateLink').href = `https://www.google.com/maps/dir/?api=1&destination=${p.lat},${p.lon}`;
  $('#description').textContent = p.description;
  $('#actualHoa').value = p.actualHoa || '';
  $('#heightChanged').value = p.heightChanged || 'No';
  $('#distance').value = p.measuredDistance || '';
  $('#status').value = p.status || 'Not started';
  $('#trimming').value = p.trimmingNotes || '';
  $('#notes').value = p.fieldNotes || '';
  $('#beforeCount').textContent = `Сохранено фото: ${p.beforePhotoCount||0}`;
  $('#afterCount').textContent = `Сохранено фото: ${p.afterPhotoCount||0}`;
  $('#poleDialog').showModal();
}

async function fileToBase64(file){
  return await new Promise((resolve,reject)=>{
    const r = new FileReader();
    r.onload = () => resolve(r.result.split(",")[1]);
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}

async function uploadSelectedPhotos(p, inputSelector, type){
  const files = [...$(inputSelector).files];
  const links = [];
  for(let i=0;i<files.length;i++){
    const f = files[i];
    const base64 = await fileToBase64(f);
    const resp = await api("uploadPhoto", {
      projectPole:p.projectPole,
      poleId:p.poleId,
      photoType:type,
      index:i+1,
      mimeType:f.type || "image/jpeg",
      base64
    });
    if(resp?.url) links.push(resp.url);
  }
  return links;
}

$('#saveBtn').addEventListener('click', async e=>{
  e.preventDefault();
  const p = poles[currentIndex];
  p.actualHoa = $('#actualHoa').value;
  p.heightChanged = $('#heightChanged').value;
  p.measuredDistance = $('#distance').value;
  p.status = $('#status').value;
  p.trimmingNotes = $('#trimming').value;
  p.fieldNotes = $('#notes').value;
  p.beforePhotoCount = Math.max(p.beforePhotoCount||0, $('#beforePhotos').files.length);
  p.afterPhotoCount = Math.max(p.afterPhotoCount||0, $('#afterPhotos').files.length);

  saveLocalState();
  render();

  try{
    const beforeLinks = await uploadSelectedPhotos(p, '#beforePhotos', 'BEFORE');
    const afterLinks = await uploadSelectedPhotos(p, '#afterPhotos', 'AFTER');
    await api("savePole", {
      projectPole:p.projectPole,
      poleId:p.poleId,
      actualHoa:p.actualHoa,
      heightChanged:p.heightChanged,
      measuredDistance:p.measuredDistance,
      trimmingNotes:p.trimmingNotes,
      status:p.status,
      fieldNotes:p.fieldNotes,
      beforeLinks,
      afterLinks
    });
  }catch(err){
    console.error(err);
    alert("Данные сохранены на телефоне, но синхронизация не удалась.");
  }
  $('#poleDialog').close();
});

$('#search').addEventListener('input', render);
$('#statusFilter').addEventListener('change', render);
$('#nearestBtn').addEventListener('click', ()=>{
  if(!navigator.geolocation) return alert('GPS недоступен');
  navigator.geolocation.getCurrentPosition(pos=>{
    const {latitude, longitude} = pos.coords;
    let best = poles[0], bestD = Infinity;
    poles.forEach(p=>{
      const d=(p.lat-latitude)**2+(p.lon-longitude)**2;
      if(d<bestD){bestD=d;best=p;}
    });
    openPole(best.projectPole);
  }, ()=>alert('Разрешите доступ к геолокации'));
});

async function boot() {
  try {
    poles = await fetch('./data.json?v=5', {
      cache: 'no-store'
    }).then(response => {
      if (!response.ok) {
        throw new Error('Ошибка загрузки data.json');
      }
      return response.json();
    });

    loadLocalState();

    // Сразу показываем столбы.
    render();

    // Google-синхронизация больше не блокирует приложение.
    if (API_URL && !API_URL.includes('PASTE_')) {
      api('getAll')
        .then(remote => {
          if (remote && remote.rows) {
            const byPole = Object.fromEntries(
              remote.rows.map(row => [Number(row.projectPole), row])
            );

            poles = poles.map(pole => ({
              ...pole,
              ...(byPole[pole.projectPole] || {})
            }));

            render();
          }
        })
        .catch(error => {
          console.warn('Работаем без синхронизации:', error);
        });
    }

  } catch (error) {
    console.error(error);

    document.querySelector('#list').innerHTML = `
      <div class="pole">
        <b>Ошибка загрузки данных</b><br>
        ${error.message}
      </div>
    `;
  }
}
boot();

if('serviceWorker' in navigator) navigator.serviceWorker.register('service-worker.js');
