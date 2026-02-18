(async function(){
    const { jsPDF } = window.jspdf;
    const $ = id => document.getElementById(id);

    // constants
    const DIV_CODES = { 'Education':'EDU','Research':'RES','Skills':'SKL','Journal Club':'JC','Guidelines':'GL','Custom':'CUS' };
    const EB_EMAILS = [
      'president-ext@cardig.me','president-int@cardig.me','vicepresident@cardig.me','treasurer@cardig.me',
      'secretary@cardig.me','caro@sci.cardig.me','caso@sci.cardig.me','caeo@sci.cardig.me','cago@sci.cardig.me',
      'cajo@sci.cardig.me','camo@sup.cardig.me','cato@sup.cardig.me','capo@sup.cardig.me','cabo@sup.cardig.me'
    ];
    const DEFAULT_LOGO_URL = 'https://cdn.prod.website-files.com/6832b0b5fc1466a06b8a6847/68a99772817d335a0c4af2cd_Round%20Logo%20(Dark).PNG';

    // DOM refs for tables/buttons
    const coordTableBody = document.querySelector('#coordTable tbody');
    const addCoord = $('addCoord'), clearCoord = $('clearCoord');
    const financeTableBody = document.querySelector('#financeTable tbody');
    const addFinance = $('addFinance'), clearFinance = $('clearFinance');
    const totalCostEl = $('totalCost');
    const sponsorTableBody = document.querySelector('#sponsorTable tbody');
    const addSponsor = $('addSponsor'), clearSponsor = $('clearSponsor');
    const requiredAmountEl = $('requiredAmount');

    // state for logo
    let logoDataUrl = '', logoAspect = 1;

    // small helper formatters
    const fmt2 = n => (Number(n)||0).toFixed(2);
    // Renaming usd to nis to handle currency formatting with NIS
    const nis = n => 'NIS ' + fmt2(n);
    const escapeHtml = s => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');

    // create a custom division input that appears when user selects "Custom"
    (function createDivisionCustomField(){
      const sel = $('division');
      const input = document.createElement('input');
      input.id = 'divisionCustom';
      input.type = 'text';
      input.placeholder = 'Custom division name';
      input.style.display = 'none';
      input.style.marginTop = '8px';
      // insert after the select (inside the same form-group)
      sel.insertAdjacentElement('afterend', input);

      // toggle visibility on change
      sel.addEventListener('change', (e) => {
        if(sel.value === 'Custom'){
          input.style.display = 'block';
          input.focus();
        } else {
          input.style.display = 'none';
        }
        updateTopRefDate();
      });
      // update ref on custom input change
      input.addEventListener('input', updateTopRefDate);
    })();

    // helper to return a user-facing division string (handles Custom)
    function getDivisionDisplay(){
      const sel = $('division');
      if(!sel) return '';
      if(sel.value === 'Custom'){
        const custom = ($('divisionCustom') && $('divisionCustom').value.trim()) || 'Custom';
        return custom;
      }
      return sel.value || '';
    }

    // load default logo into preview and a dataURL for PDF
    async function fetchImageToDataURL(url){
      try{
        const resp = await fetch(url, {mode:'cors'});
        if(!resp.ok) throw new Error('Image fetch failed ' + resp.status);
        const blob = await resp.blob();
        return await new Promise((resolve,reject)=>{
          const r = new FileReader();
          r.onload = () => {
            const dataUrl = r.result;
            const img = new Image();
            img.onload = () => resolve({ dataUrl, aspect: img.height / img.width || 1 });
            img.onerror = () => resolve({ dataUrl, aspect: 1 });
            img.src = dataUrl;
          };
          r.onerror = reject;
          r.readAsDataURL(blob);
        });
      }catch(e){ console.warn(e); return null; }
    }
    async function loadDefaultLogoIfChecked(){
      try{
        const use = $('useDefaultLogo') && $('useDefaultLogo').checked;
        if(!use){ logoDataUrl = ''; document.getElementById('logoPreview').src = ''; return; }
        const res = await fetchImageToDataURL(DEFAULT_LOGO_URL);
        if(res){ logoDataUrl = res.dataUrl; logoAspect = res.aspect; document.getElementById('logoPreview').src = logoDataUrl; }
        else { logoDataUrl = ''; document.getElementById('logoPreview').src = DEFAULT_LOGO_URL; }
      }catch(e){ console.warn('logo load', e); logoDataUrl = ''; }
    }

    // coordinator helpers (Officer email selection)
    function createEmailControlForRole(role, value=''){
      if(role === 'Officer'){
        const select = document.createElement('select');
        select.className = 'coord-email-select';
        const emptyOpt = document.createElement('option'); emptyOpt.value=''; emptyOpt.textContent='Choose Email'; select.appendChild(emptyOpt);
        EB_EMAILS.forEach(e=>{ const o = document.createElement('option'); o.value=e; o.textContent=e; if(e===value) o.selected=true; select.appendChild(o); });
        return select;
      } else {
        const input = document.createElement('input'); input.type = 'email'; input.className='coord-email-input'; input.placeholder='Email'; input.value = value || ''; return input;
      }
    }
    function swapEmailControl(tr, role){
      const cell = tr.querySelector('.coord-email-cell');
      const old = cell.querySelector('.coord-email-select, .coord-email-input');
      let prevVal = '';
      if(old) prevVal = old.value || '';
      cell.innerHTML = '';
      const ctrl = createEmailControlForRole(role, prevVal);
      cell.appendChild(ctrl);
    }
    function addCoordinatorRow(name='', email='', whatsapp='', position='', role='Officer'){
      const tr = document.createElement('tr');
      tr.innerHTML = `<td><input class="coord-name" type="text" value="${escapeHtml(name)}" placeholder="Name"></td>
                      <td class="coord-email-cell"></td>
                      <td><input class="coord-wh" type="text" value="${escapeHtml(whatsapp)}" placeholder="Phone"></td>
                      <td><input class="coord-pos" type="text" value="${escapeHtml(position)}" placeholder="Pos"></td>
                      <td>
                        <select class="coord-role">
                          <option${role==='Officer'? ' selected':''}>Officer</option>
                          <option${role==='Assistant'? ' selected':''}>Assistant</option>
                          <option${role==='Volunteer'? ' selected':''}>Volunteer</option>
                        </select>
                      </td>
                      <td style="text-align:center"><button type="button" class="del-coord btn-ghost">Del</button></td>`;
      coordTableBody.appendChild(tr);
      swapEmailControl(tr, role);
      if(role === 'Officer' && email){
        const sel = tr.querySelector('.coord-email-select');
        if(sel){ const opt = Array.from(sel.options).find(o=>o.value===email); if(opt) opt.selected=true; }
      } else {
        const input = tr.querySelector('.coord-email-input');
        if(input) input.value = email || '';
      }
      tr.querySelector('.coord-role').addEventListener('change', e => swapEmailControl(tr, e.target.value));
      tr.querySelector('.del-coord').addEventListener('click', ()=> tr.remove());
    }
    addCoord.addEventListener('click', ()=> addCoordinatorRow());
    clearCoord.addEventListener('click', ()=> { coordTableBody.innerHTML=''; });

    // finance rows
    function parseNumberSafe(raw){
      // robust parse: accept strings like "1,234.56" and empty
      try{
        if(raw === undefined || raw === null) return 0;
        const s = String(raw).trim().replace(/,/g,'');
        const v = parseFloat(s);
        return isNaN(v) ? 0 : v;
      }catch(e){ return 0; }
    }
    function addFinanceRow(item='', price=0, qty=1){
      const tr = document.createElement('tr');
      tr.innerHTML = `<td><input class="fin-item" type="text" value="${escapeHtml(item)}" placeholder="Item"></td>
                      <td><input class="fin-price" type="number" step="0.01" min="0" value="${fmt2(price)}"></td>
                      <td><input class="fin-qty" type="number" step="1" min="0" value="${qty}"></td>
                      <td class="fin-total" style="font-weight:600">${fmt2(price*qty)}</td>
                      <td style="text-align:center"><button type="button" class="del-fin btn-ghost">Del</button></td>`;
      financeTableBody.appendChild(tr);
      // use delegated input handling as well as per-input listeners
      const priceInput = tr.querySelector('.fin-price');
      const qtyInput = tr.querySelector('.fin-qty');
      [priceInput, qtyInput].forEach(i=> {
        if(i) i.addEventListener('input', ()=>{ recalcFinance(); updateRequiredAmount(); updateTopRefDate(); });
      });
      tr.querySelector('.del-fin').addEventListener('click', ()=> { tr.remove(); recalcFinance(); updateRequiredAmount(); updateTopRefDate(); });
      recalcFinance();
    }
    addFinance.addEventListener('click', ()=> addFinanceRow());
    clearFinance.addEventListener('click', ()=> { financeTableBody.innerHTML=''; recalcFinance(); updateRequiredAmount(); updateTopRefDate(); });

    function recalcFinance(){
      // robust numeric parsing and update of row totals + footer
      let total = 0;
      const rows = document.querySelectorAll('#financeTable tbody tr');
      rows.forEach(tr=>{
        const priceRaw = (tr.querySelector('.fin-price')||{}).value;
        const qtyRaw = (tr.querySelector('.fin-qty')||{}).value;
        const price = parseNumberSafe(priceRaw);
        const qty = parseNumberSafe(qtyRaw);
        const amt = Math.round(price * qty * 100)/100;
        const td = tr.querySelector('.fin-total');
        if(td) td.textContent = fmt2(amt);
        total += amt;
      });
      if(totalCostEl) totalCostEl.textContent = fmt2(total);
      return total;
    }

    // sponsors
    function addSponsorRow(name='', totalFunding=0){
      const tr = document.createElement('tr');
      tr.innerHTML = `<td><input class="sponsor-name" type="text" value="${escapeHtml(name)}" placeholder="Sponsor"></td>
                      <td><input class="sponsor-total" type="number" step="0.01" min="0" value="${fmt2(totalFunding)}"></td>
                      <td style="text-align:center"><button type="button" class="del-sponsor btn-ghost">Del</button></td>`;
      sponsorTableBody.appendChild(tr);
      tr.querySelector('.del-sponsor').addEventListener('click', ()=> { tr.remove(); updateRequiredAmount(); updateTopRefDate(); });
      const totalInput = tr.querySelector('.sponsor-total');
      if(totalInput) totalInput.addEventListener('input', () => { updateRequiredAmount(); updateTopRefDate(); });
      updateRequiredAmount();
    }
    addSponsor.addEventListener('click', ()=> addSponsorRow());
    clearSponsor.addEventListener('click', ()=> { sponsorTableBody.innerHTML=''; updateRequiredAmount(); updateTopRefDate(); });

    // compute required amount
    function updateRequiredAmount(){
      // ensure finance totals up-to-date first
      const totalCost = recalcFinance();
      let sponsorSum = 0;
      document.querySelectorAll('#sponsorTable tbody tr').forEach(tr=>{
        const vRaw = (tr.querySelector('.sponsor-total')||{}).value;
        const v = parseNumberSafe(vRaw);
        sponsorSum += v;
      });
      const required = Math.max(0, Math.round((totalCost - sponsorSum) * 100)/100);
      if(requiredAmountEl) requiredAmountEl.textContent = fmt2(required);
      return required;
    }

    // initial seed
    (function bootstrap(){
      if(coordTableBody.children.length === 0) addCoordinatorRow('Coordinator Name','vicepresident@cardig.me','+970-59-0000000','', 'Officer');
      if(financeTableBody.children.length === 0) addFinanceRow('Venue',0,0);
      if(sponsorTableBody.children.length === 0) addSponsorRow('',0);
      updateRequiredAmount();
    })();
    
    // observe changes to finance and sponsor tbody (rows being added/removed) and recalc
    (function attachTableObservers(){
      const onFinanceMut = (mutList) => { recalcFinance(); updateRequiredAmount(); updateTopRefDate(); };
      const observerOpts = { childList: true, subtree: false };
      const fObs = new MutationObserver(onFinanceMut);
      const sObs = new MutationObserver(onFinanceMut);
      if(financeTableBody) fObs.observe(financeTableBody, observerOpts);
      if(sponsorTableBody) sObs.observe(sponsorTableBody, observerOpts);
      // also delegate input events to ensure any dynamic inputs are picked up
      financeTableBody.addEventListener('input', (e)=> { if(e.target.matches('.fin-price, .fin-qty')) { recalcFinance(); updateRequiredAmount(); } });
      sponsorTableBody.addEventListener('input', (e)=> { if(e.target.matches('.sponsor-total')) { updateRequiredAmount(); } });
    })();

    // ref generator
    function genRef(){
      const raw = ($('activityName').value || 'activity').replace(/\s+/g,'_').slice(0,30);
      const name = raw.replace(/[^a-zA-Z0-9_-]/g,'');
      const now = new Date();
      const date = now.toISOString().slice(0,10);
      const divSel = ($('division').value || '').trim();
      const divKey = divSel === 'Custom' ? 'Custom' : divSel;
      const code = DIV_CODES[divKey] || 'GEN';
      return `ACT-${date}-${code}-${name}`;
    }

    // update header Ref and Date
    function updateTopRefDate(){
      const ref = genRef();
      const today = new Date().toLocaleDateString();
      const topRefEl = $('topRef');
      const topDateEl = $('topDate');
      if(topRefEl) topRefEl.textContent = ref; // simplified text for new badge
      if(topDateEl) topDateEl.textContent = today;
    }

    // ensure vertical space (add page when needed)
    function ensureSpace(doc, yRef, needed, margin, footerReserve=14){
      const pageH = doc.internal.pageSize.getHeight();
      if(yRef + needed > pageH - margin - footerReserve){
        doc.addPage();
        return margin + 15; // Return new Y at top of fresh page
      }
      return yRef;
    }

    // --- MODERN REDESIGN: BUILD PDF FUNCTION (FIXED JERUSALEM HEADER) ---
    function buildPDF(){
      const doc = new jsPDF({unit:'mm', format:'a4'});
      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();
      const margin = 12; // Standard margin
      let y = margin;

      // Colors - Clinical/Professional Brand
      const colPrimary = [9, 41, 82];     // #092952 (Brand Blue)
      // Updated Accent to match Primary (#092952) for headers
      const colAccent = [9, 41, 82];      
      const colGray = [80, 80, 80];       // Dark Gray text
      const colLightBg = [240, 244, 248]; // Very light blue-gray
      const colLine = [200, 200, 200];    // Divider lines

      // --- 1. HEADER SECTION ---
      // Logo (Left)
      let logoH = 20;
      if(logoDataUrl){
        try{
          const imgType = logoDataUrl.startsWith('data:image/png') ? 'PNG' : 'JPEG';
          const logoW = 20;
          logoH = logoW * (logoAspect || 1);
          doc.addImage(logoDataUrl, imgType, margin, y, logoW, logoH);
        }catch(e){ console.warn('logo fail', e); }
      }

      // Institution Name - SPLIT TO PREVENT CUTOFF
      const headerTextX = margin + 25; 
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(14);
      doc.setTextColor(...colPrimary);
      doc.text("CARDIOLOGY INTEREST GROUP", headerTextX, y + 5);
      doc.text("OF JERUSALEM", headerTextX, y + 11);

      // UPDATED ADDRESS
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.setTextColor(...colGray);
      doc.text("Faculty of Medicine, Al-Quds University", headerTextX, y + 16);
      doc.text("P.O. Box 20002, Abu Dis, Jerusalem, Palestine", headerTextX, y + 20);

      // Document Title Badge (Top Right)
      doc.setFillColor(...colPrimary);
      doc.roundedRect(pageW - margin - 55, y, 55, 9, 1, 1, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(9);
      doc.setFont('helvetica', 'bold');
      doc.text("ACTIVITY ENROLLMENT FORM", pageW - margin - 27.5, y + 6, {align:'center'});

      // Reference & Date
      const ref = genRef();
      doc.setTextColor(...colGray);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.text(`Ref ID: ${ref}`, pageW - margin, y + 15, {align:'right'});
      doc.text(`Date: ${new Date().toLocaleDateString()}`, pageW - margin, y + 19, {align:'right'});

      // Adjusted Y to account for extra address line
      y += Math.max(logoH, 26) + 6;
      
      // Thick Divider
      doc.setDrawColor(...colPrimary);
      doc.setLineWidth(0.8);
      doc.line(margin, y, pageW - margin, y);
      y += 8;

      // --- 2. DEMOGRAPHICS BOX (FIXED DATA) ---
      // Structured grid background
      const boxHeight = 44; 
      doc.setFillColor(...colLightBg);
      doc.setDrawColor(...colLine);
      doc.setLineWidth(0.1);
      doc.rect(margin, y, pageW - 2*margin, boxHeight, 'FD');

      // Fetch Form Data
      const actName = ($('activityName').value || '—').toUpperCase();
      const divVal = getDivisionDisplay() || '—';
      const dateStr = $('actDate').value ? new Date($('actDate').value).toLocaleDateString() : '—';
      const timeStr = `${$('timeFrom').value || '—'} - ${$('timeTo').value || '—'}`;
      const place = $('place').value || '—';
      const coordName = $('coordinatorName').value || '—';
      const coordPos = $('coordinatorPosition').value || '';
      const coordFull = coordName + (coordPos ? ` (${coordPos})` : '');
      const targetGroup = $('targetGroup').value || '—';
      const expectedNum = $('expectedNum').value || '—';
      
      // Sub Activity Logic
      const isSub = $('isSubActivity').value === 'yes';
      const primaryAct = isSub ? ($('primaryActivityName').value || 'Unknown Parent') : null;

      // Rows Configuration
      const row1Lab = y + 5;  const row1Val = y + 11;
      const row2Lab = y + 17; const row2Val = y + 23;
      const row3Lab = y + 29; const row3Val = y + 35;
      
      // Row 1: Title, Division, Date
      doc.setTextColor(...colGray); doc.setFont('helvetica', 'bold'); doc.setFontSize(7);
      doc.text("ACTIVITY TITLE", margin + 4, row1Lab);
      doc.text("DIVISION", margin + 80, row1Lab);
      doc.text("DATE / TIME", margin + 130, row1Lab);

      doc.setTextColor(0,0,0); doc.setFont('helvetica', 'bold'); doc.setFontSize(10);
      doc.text(actName, margin + 4, row1Val);
      doc.setFont('helvetica', 'normal');
      doc.text(divVal, margin + 80, row1Val);
      doc.text(`${dateStr}  |  ${timeStr}`, margin + 130, row1Val);

      // Row 2: Coordinator, Location
      doc.setTextColor(...colGray); doc.setFont('helvetica', 'bold'); doc.setFontSize(7);
      doc.text("COORDINATOR", margin + 4, row2Lab);
      doc.text("LOCATION / PLATFORM", margin + 80, row2Lab);

      doc.setTextColor(0,0,0); doc.setFont('helvetica', 'normal'); doc.setFontSize(10);
      doc.text(coordFull, margin + 4, row2Val);
      doc.text(place, margin + 80, row2Val);

      // Row 3: Target, Expected, Sub-Activity
      doc.setTextColor(...colGray); doc.setFont('helvetica', 'bold'); doc.setFontSize(7);
      doc.text("TARGET GROUP", margin + 4, row3Lab);
      doc.text("EXPECTED PARTICIPANTS", margin + 80, row3Lab);
      if(primaryAct) doc.text("SUB-ACTIVITY OF", margin + 130, row3Lab);

      doc.setTextColor(0,0,0); doc.setFont('helvetica', 'normal'); doc.setFontSize(10);
      doc.text(targetGroup, margin + 4, row3Val);
      doc.text(expectedNum, margin + 80, row3Val);
      if(primaryAct) {
        doc.setFont('helvetica', 'italic');
        doc.text(primaryAct, margin + 130, row3Val);
      }

      y += boxHeight + 8;

      // --- 3. CLINICAL/REPORT SECTIONS ---
      function addSectionHeader(title, yPos) {
        doc.setFillColor(...colAccent); // Using the updated #092952 color
        doc.rect(margin, yPos, pageW - 2*margin, 6, 'F');
        doc.setTextColor(255, 255, 255);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(9);
        doc.text(title.toUpperCase(), margin + 2, yPos + 4.2);
        return yPos + 10;
      }

      function addTextParam(label, text, yPos) {
        doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(...colPrimary);
        const labelW = doc.getTextWidth(label + ": ");
        doc.text(label + ":", margin, yPos);
        
        doc.setFont('helvetica', 'normal'); doc.setTextColor(0);
        const textW = pageW - 2*margin - labelW - 2;
        const safeText = (text && text.trim()) ? text.trim() : '—';
        const lines = doc.splitTextToSize(safeText, textW);
        doc.text(lines, margin + labelW + 2, yPos);
        return yPos + (lines.length * 5) + 3;
      }

      // Proposal
      y = addSectionHeader("Proposal & Objectives", y);
      
      const selCrit = $('selectionCriteria').value.trim();
      if(selCrit) {
        y = addTextParam("Selection Criteria", selCrit, y);
        y += 2;
      }
      
      y = addTextParam("Purpose", $('purpose').value, y);
      y = addTextParam("Objectives", $('objectives').value, y);
      y = addTextParam("Methodology", $('methodology').value, y);
      y += 4;

      // --- 4. DATA TABLES (Grid Style) ---
      
      // Coordinators
      y = ensureSpace(doc, y, 40, margin);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...colPrimary);
      doc.text("COORDINATION TEAM", margin, y);
      doc.setDrawColor(...colLine); doc.setLineWidth(0.2);
      doc.line(margin, y+1, pageW-margin, y+1);
      y += 5;

      const coordRows = [];
      document.querySelectorAll('#coordTable tbody tr').forEach(tr=>{
        const nm = (tr.querySelector('.coord-name')||{}).value;
        const role = (tr.querySelector('.coord-role')||{}).value;
        const pos = (tr.querySelector('.coord-pos')||{}).value;
        const wh = (tr.querySelector('.coord-wh')||{}).value;
        const sel = tr.querySelector('.coord-email-select');
        const inp = tr.querySelector('.coord-email-input');
        const em = (sel && sel.value) ? sel.value : (inp && inp.value ? inp.value : '');
        coordRows.push([nm||'—', role||'—', pos||'—', em||'—', wh||'—']);
      });

      doc.autoTable({
        startY: y,
        head: [['Name', 'Role', 'Position', 'Email', 'Phone']],
        body: coordRows.length ? coordRows : [['—','—','—','—','—']],
        theme: 'grid',
        styles: { fontSize: 8, cellPadding: 3, textColor: [20,20,20], lineColor: [180,180,180], lineWidth: 0.1 },
        headStyles: { fillColor: [240, 240, 240], textColor: [0,0,0], fontStyle: 'bold' },
        margin: { left: margin, right: margin }
      });
      y = doc.lastAutoTable.finalY + 10;

      // Financials
      y = ensureSpace(doc, y, 40, margin);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...colPrimary);
      doc.text("FINANCIAL BREAKDOWN", margin, y);
      doc.line(margin, y+1, pageW-margin, y+1);
      y += 5;

      const financeRows = [];
      document.querySelectorAll('#financeTable tbody tr').forEach(tr=>{
        const item = (tr.querySelector('.fin-item')||{}).value;
        const price = parseNumberSafe((tr.querySelector('.fin-price')||{}).value);
        const qty = parseNumberSafe((tr.querySelector('.fin-qty')||{}).value);
        const total = parseNumberSafe((tr.querySelector('.fin-total')||{}).textContent);
        financeRows.push([item||'—', qty, nis(price), nis(total)]);
      });
      const totalCost = recalcFinance();

      doc.autoTable({
        startY: y,
        head: [['Item', 'Qty', 'Unit Price', 'Total']],
        body: financeRows.length ? financeRows : [['No items','','','0.00']],
        theme: 'grid',
        styles: { fontSize: 8, cellPadding: 3, textColor: [20,20,20], lineColor: [180,180,180], lineWidth: 0.1 },
        headStyles: { fillColor: [240, 240, 240], textColor: [0,0,0], fontStyle: 'bold' },
        columnStyles: { 1:{halign:'center'}, 2:{halign:'right'}, 3:{halign:'right'} },
        margin: { left: margin, right: margin },
        foot: [['', '', 'Total Est. Cost', nis(totalCost)]],
        footStyles: { fillColor: [255,255,255], textColor: colPrimary, fontStyle:'bold', halign:'right' }
      });
      y = doc.lastAutoTable.finalY + 6;

      // Sponsors Summary
      const sponsorRows = [];
      let sponsorSum = 0;
      document.querySelectorAll('#sponsorTable tbody tr').forEach(tr=>{
        const name = (tr.querySelector('.sponsor-name')||{}).value;
        if(name){
           const val = parseNumberSafe((tr.querySelector('.sponsor-total')||{}).value);
           sponsorSum += val;
           sponsorRows.push(`${name} (${nis(val)})`);
        }
      });
      const requiredTotal = Math.max(0, Math.round((totalCost - sponsorSum) * 100)/100);

      doc.setFont('helvetica', 'normal'); doc.setFontSize(8);
      if(sponsorRows.length) {
        doc.text("Funding Sources: " + sponsorRows.join(', '), margin, y);
        y += 5;
      }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...colPrimary);
      doc.text(`NET FUNDS REQUIRED: ${nis(requiredTotal)}`, margin, y);
      y += 12;

      // --- 5. LOGISTICS & APPROVALS (VISUAL 3-COLUMN) ---
      y = ensureSpace(doc, y, 70, margin);
      y = addSectionHeader("Logistics, Branding & Admin Requests", y);

      // Create distinct blocks for requests (3 columns visually)
      const colW = (pageW - 2*margin) / 3;
      
      // COL 1: Logistics (CaPO)
      let curY = y;
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(...colPrimary);
      doc.text("Logistics (CaPO)", margin, curY); curY += 5;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(0);
      
      const capoItems = [];
      if($('reqProjector').checked) capoItems.push("• Projector");
      if($('reqPrints').checked) capoItems.push("• Prints");
      if($('reqHall').checked) capoItems.push("• Hall Reservation");
      const noteCapo = $('customCaPO').value.trim();
      
      if(capoItems.length === 0 && !noteCapo) doc.text("— None", margin, curY);
      else {
        capoItems.forEach(i => { doc.text(i, margin, curY); curY+=4; });
        if(noteCapo) {
            const lines = doc.splitTextToSize("Note: " + noteCapo, colW - 4);
            doc.text(lines, margin, curY);
            curY += lines.length*4;
        }
      }
      let maxY = curY;

      // COL 2: Branding (CaBO)
      curY = y;
      const x2 = margin + colW;
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(...colPrimary);
      doc.text("Media & Branding (CaBO)", x2, curY); curY += 5;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(0);
      
      const caboItems = [];
      if($('reqAnnounce').checked) caboItems.push("• Announcement");
      if($('reqJoint').checked) caboItems.push("• Joint Activity");
      if($('reqDesign').value === 'yes') caboItems.push("• Design Needed");
      const noteCabo = $('customCaBO').value.trim();
      
      if(caboItems.length === 0 && !noteCabo) doc.text("— None", x2, curY);
      else {
        caboItems.forEach(i => { doc.text(i, x2, curY); curY+=4; });
        if(noteCabo) {
            const lines = doc.splitTextToSize("Note: " + noteCabo, colW - 4);
            doc.text(lines, x2, curY);
            curY += lines.length*4;
        }
      }
      if(curY > maxY) maxY = curY;

      // COL 3: Admin (Secretary)
      curY = y;
      const x3 = margin + colW*2;
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(...colPrimary);
      doc.text("Secretary (Docs)", x3, curY); curY += 5;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(0);
      
      const noteDocs = $('docsRequested').value.trim();
      if(!noteDocs) doc.text("— None", x3, curY);
      else {
          const lines = doc.splitTextToSize(noteDocs, colW - 4);
          doc.text(lines, x3, curY);
          curY += lines.length*4;
      }
      if(curY > maxY) maxY = curY;
      
      y = maxY + 8;


      // SIGNATURE BLOCK
      y = ensureSpace(doc, y, 40, margin);
      const approvalBody = [
        ['Internal President', '', ''],
        ['External President', '', ''],
        ['Vice President', '', ''],
        ['Secretary', '', ''],
        ['Treasurer', '', '']
      ];

      doc.autoTable({
        startY: y,
        head: [['Executive Role', 'Signature / Stamp', 'Date']],
        body: approvalBody,
        theme: 'grid',
        styles: { fontSize: 9, cellPadding: 4, minCellHeight: 12, lineColor: [150,150,150], lineWidth: 0.1 },
        headStyles: { fillColor: [230, 235, 240], textColor: [0,0,0], fontStyle:'bold' },
        columnStyles: { 0: {cellWidth: 60, fontStyle:'bold'}, 2: {cellWidth: 35} },
        margin: { left: margin, right: margin }
      });

      // Footer
      const pageCount = doc.internal.getNumberOfPages();
      for(let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8);
        doc.setTextColor(150);
        doc.text(`CARDIG Activity Enrollment | Ref: ${ref} | Page ${i} of ${pageCount}`, pageW / 2, pageH - 8, { align: 'center' });
      }

      return { doc, ref };
    }

    // download the main PDF (new tab)
    async function downloadMainPDF(){
      try{
        await loadDefaultLogoIfChecked();
        const { doc, ref } = buildPDF();
        const blob = doc.output('blob'); const url = URL.createObjectURL(blob); const filename = `${ref}.pdf`;
        const newWin = window.open('', '_blank');
        if(newWin){
          const safeHtml = `<!doctype html><html><head><meta charset="utf-8"><title>Downloading ${filename}</title></head>
            <body><a id="dl" href="${url}" download="${filename}">Download</a>
            <scr` + `ipt>const a=document.getElementById('dl');a.click();setTimeout(()=>{try{window.close()}catch(e){}},1500);</scr` + `ipt>
            </body></html>`;
          newWin.document.open(); newWin.document.write(safeHtml); newWin.document.close();
        } else {
          const a = document.createElement('a'); a.href = url; a.download = filename; a.target = '_blank';
          document.body.appendChild(a); a.click(); a.remove();
        }
        setTimeout(()=> URL.revokeObjectURL(url), 20000);
      }catch(e){ console.error(e); alert('Could not create or download PDF — check console.'); }
    }

    // Export a compact "card"
    async function exportCard(type){
      await loadDefaultLogoIfChecked();
      const doc = new jsPDF({unit:'mm', format:'a6'});
      const w = doc.internal.pageSize.getWidth();
      const h = doc.internal.pageSize.getHeight();
      const margin = 8;
      let y = margin;

      // header band
      doc.setFillColor(9,41,82);
      doc.rect(0, 0, w, 24, 'F');

      if(logoDataUrl){
        try{
          const imgType = logoDataUrl.startsWith('data:image/png') ? 'PNG' : 'JPEG';
          doc.addImage(logoDataUrl, imgType, margin, 3, 18, 18 * (logoAspect||1));
        }catch(e){ console.warn('card logo fail', e); }
      }

      doc.setFont('helvetica','bold'); doc.setFontSize(12); doc.setTextColor(255,255,255);
      doc.text(type, w - margin, 10, {align:'right'});

      y = 32;

      doc.setFont('helvetica','bold'); doc.setFontSize(12); doc.setTextColor(9,41,82);
      const actTitle = ($('activityName').value || '—');
      const titleLines = doc.splitTextToSize(actTitle, w - margin*2);
      doc.text(titleLines, margin, y);
      y += titleLines.length * 6 + 4;

      const metaPairs = [
        { h: 'Division', b: getDivisionDisplay() || '—' },
        { h: 'Date', b: ($('actDate').value ? new Date($('actDate').value).toLocaleDateString() : '—') },
        { h: 'Time', b: `${$('timeFrom').value || '—'} - ${$('timeTo').value || '—'}` },
        { h: 'Place', b: $('place').value || '—' },
        { h: 'Coordinator', b: $('coordinatorName').value || '—' }
      ];

      for(const p of metaPairs){
        doc.setFont('helvetica','bold'); doc.setFontSize(9.5); doc.setTextColor(9,41,82);
        doc.text(`${p.h}:`, margin, y);
        const headerWidth = doc.getTextWidth(p.h + ': ') + 4;
        doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(50);
        const lines = doc.splitTextToSize(p.b, w - margin*2 - headerWidth);
        doc.text(lines, margin + headerWidth, y);
        y += Math.max(lines.length,1) * 5 + 3;
      }
      y += 2;

      if(type === 'Financial Request'){
        doc.setFont('helvetica','bold'); doc.setFontSize(10); doc.setTextColor(9,41,82);
        doc.text('Financial Items', margin, y); y += 6;
        doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(50);

        const rows = [];
        document.querySelectorAll('#financeTable tbody tr').forEach(tr=>{
          const item = (tr.querySelector('.fin-item')||{}).value || '';
          const price = parseNumberSafe(((tr.querySelector('.fin-price')||{}).value||0));
          const qty = parseNumberSafe(((tr.querySelector('.fin-qty')||{}).value||0));
          const amt = Math.round(price * qty * 100)/100;
          rows.push({ item: item || '—', detail: `${qty} x ${fmt2(price)} = ${fmt2(amt)}`, amt });
        });
        if(rows.length === 0) rows.push({ item: 'No items', detail: '', amt: 0 });

        rows.forEach(r=>{
          doc.setFont('helvetica','bold'); doc.setFontSize(9); doc.text(r.item, margin, y);
          const left = doc.getTextWidth(r.item) + 6;
          doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.text(r.detail, margin + left, y);
          y += 5.5;
        });

        y += 4;
        doc.setFont('helvetica','bold'); doc.setFontSize(10);
        doc.text(`Total: ${nis(recalcFinance())}`, margin, y);

        const required = updateRequiredAmount();
        y += 6;
        doc.setFont('helvetica','bold'); doc.setFontSize(10); doc.text(`Required: ${nis(required)}`, margin, y);
      }

      if(type === 'Request for CaBO'){
        doc.setFont('helvetica','bold'); doc.setFontSize(10); doc.setTextColor(9,41,82);
        doc.text('CaBO Requests', margin, y); y += 6;
        doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(50);
        const ceboReqs = [
          `Announce: ${$('reqAnnounce').checked ? 'YES' : 'NO'}`,
          `Joint: ${$('reqJoint').checked ? 'YES' : 'NO'}`,
          `Joint org: ${$('jointOrgName').value || '—'}`,
          `Designs: ${$('reqDesign').value === 'yes' ? 'YES' : 'NO'}`
        ].join(' • ');
        doc.text(doc.splitTextToSize(ceboReqs, w - margin*2), margin, y); y += 8;
        doc.setFont('helvetica','italic'); doc.setFontSize(9); doc.text('Custom requests:', margin, y); y += 6;
        doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.text(doc.splitTextToSize($('customCaBO').value.trim() || 'None', w - margin*2), margin, y);
      }

      if(type === 'Request for CaPO'){
        doc.setFont('helvetica','bold'); doc.setFontSize(10); doc.setTextColor(9,41,82);
        doc.text('CaPO Requests', margin, y); y += 6;
        doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(50);
        const extraReqs = [
          `Projector: ${$('reqProjector').checked ? 'YES' : 'NO'}`,
          `Prints: ${$('reqPrints').checked ? 'YES' : 'NO'}`,
          `Hall: ${$('reqHall').checked ? 'YES' : 'NO'}`
        ].join(' • ');
        doc.text(doc.splitTextToSize(extraReqs, w - margin*2), margin, y); y += 8;
        doc.setFont('helvetica','italic'); doc.setFontSize(9); doc.text('Custom requests:', margin, y); y += 6;
        doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.text(doc.splitTextToSize($('customCaPO').value.trim() || 'None', w - margin*2), margin, y);
      }

      if(type === 'Request for Secretary'){
        doc.setFont('helvetica','bold'); doc.setFontSize(10); doc.setTextColor(9,41,82);
        doc.text('Documents Requested', margin, y); y += 6;
        doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.text(doc.splitTextToSize($('docsRequested').value.trim() || 'None', w - margin*2), margin, y);
      }

      doc.setFont('helvetica','normal'); doc.setFontSize(7.5); doc.setTextColor(110);
      doc.text(genRef(), w - margin, h - 6, {align:'right'});

      const blob = doc.output('blob'); const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = `${type.replace(/\s+/g,'_')}_${genRef()}.pdf`; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(()=> URL.revokeObjectURL(url), 10000);
    }

    $('exportFinance').addEventListener('click', ()=> exportCard('Financial Request'));
    $('exportCaBO').addEventListener('click', ()=> exportCard('Request for CaBO'));
    $('exportCaPO').addEventListener('click', ()=> exportCard('Request for CaPO'));
    $('exportSec').addEventListener('click', ()=> exportCard('Request for Secretary'));
    $('downloadBtn').addEventListener('click', async (e)=> { e.preventDefault(); await downloadMainPDF(); });
    $('useDefaultLogo').addEventListener('change', ()=> loadDefaultLogoIfChecked());
    await loadDefaultLogoIfChecked();

    (function populateNotesAndChecklist(){
      const notesEl = $('notesText');
      notesEl.innerHTML = '';
      const p = document.createElement('div');
      p.innerHTML = `This form must be submitted to the Vice President (<strong>vicepresident@cardig.me</strong>) at least <strong>10 days</strong> prior to any on-ground activity and at least <strong>4 days</strong> prior to online activities. The activity will be reviewed and approved by the Board within a maximum of <strong>3 days</strong> if it is on-ground and requires financial approval, within <strong>2 days</strong> if it is on-ground and does not require financial approval, and within <strong>1 day</strong> if it is an online activity.`;
      notesEl.appendChild(p);

      const checklistEl = $('checklist');
      checklistEl.innerHTML = '';
      const bullets = [
        'Complete activity details & proposal',
        'Add coordinators (set role; choose official email for Officers)',
        'List finance items & sponsors if funding needed',
        'Use CaPO/CaBO custom requests for special needs',
        'Send form to VP early: see deadlines above'
      ];
      bullets.forEach(b => {
        const li = document.createElement('li'); li.textContent = b; checklistEl.appendChild(li);
      });
    })();

    $('activityName').addEventListener('input', updateTopRefDate);
    $('division').addEventListener('change', updateTopRefDate);
    if($('divisionCustom')) $('divisionCustom').addEventListener('input', updateTopRefDate);
    $('actDate').addEventListener('change', updateTopRefDate);

    updateTopRefDate();
  })();
