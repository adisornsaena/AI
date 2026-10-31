// ช่องเลือกโมเดลที่ใช้ร่วมกันทุกหน้า: ใส่ <select id="model"></select> แล้วเรียก getModel() ตอนส่งคำขอ
(function () {
  const sel = document.getElementById("model");
  const KEY = "okmd-demo-model";
  const read = () => { try { return localStorage.getItem(KEY); } catch (e) { return null; } };
  const save = (v) => { try { localStorage.setItem(KEY, v); } catch (e) { /* ไม่เป็นไร: แค่จำค่าไม่ได้ */ } };

  window.getModel = () => sel.value || "";

  fetch("/api/models")
    .then((r) => r.json())
    .then((data) => {
      if (!data.models) throw new Error(data.error || "no models");
      const saved = read();
      sel.innerHTML = "";
      data.models.forEach((m) => {
        const o = document.createElement("option");
        o.value = o.textContent = m;
        sel.appendChild(o);
      });
      sel.value = data.models.includes(saved) ? saved : data.default;
    })
    .catch(() => {
      sel.innerHTML = '<option value="">(ค่าเริ่มต้น)</option>';
    });

  sel.addEventListener("change", () => save(sel.value));
})();
