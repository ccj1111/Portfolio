(function () {
  const itemsRows = document.getElementById("itemsRows");
  const toppingsRows = document.getElementById("toppingsRows");
  if (!itemsRows || !toppingsRows) return; // not on the admin menu page

  const initial = window.__INITIAL_MENU__ || { items: [], toppings: [] };
  const templates = JSON.parse(document.getElementById("templatesData").textContent || "[]");

  function makeRow(container, name, price) {
    const row = document.createElement("div");
    row.className = "row-flex";
    row.innerHTML = `
      <input type="text" class="row-name" placeholder="名稱" value="${escapeAttr(name || "")}">
      <input type="number" class="row-price" placeholder="價格" min="0" step="1" value="${price ?? ""}" style="max-width:110px;">
      <button type="button" class="btn secondary row-remove">刪除</button>
    `;
    row.querySelector(".row-remove").addEventListener("click", () => row.remove());
    container.appendChild(row);
  }

  function escapeAttr(str) {
    return String(str).replace(/"/g, "&quot;");
  }

  function collectRows(container) {
    return Array.from(container.querySelectorAll(".row-flex"))
      .map((row) => ({
        name: row.querySelector(".row-name").value.trim(),
        price: parseInt(row.querySelector(".row-price").value, 10),
      }))
      .filter((r) => r.name && !Number.isNaN(r.price));
  }

  function clearRows(container) {
    container.innerHTML = "";
  }

  // Initial render
  if (initial.items.length) {
    initial.items.forEach((i) => makeRow(itemsRows, i.name, i.price));
  } else {
    makeRow(itemsRows, "", "");
  }
  if (initial.toppings.length) {
    initial.toppings.forEach((t) => makeRow(toppingsRows, t.name, t.price));
  }

  document.getElementById("addItemBtn").addEventListener("click", () => makeRow(itemsRows, "", ""));
  document.getElementById("addToppingBtn").addEventListener("click", () => makeRow(toppingsRows, "", ""));

  // Sync hidden JSON fields right before submit
  const menuForm = document.getElementById("menuForm");
  menuForm.addEventListener("submit", () => {
    document.getElementById("itemsJson").value = JSON.stringify(collectRows(itemsRows));
    document.getElementById("toppingsJson").value = JSON.stringify(collectRows(toppingsRows));
  });

  // Apply template
  const templateSelect = document.getElementById("templateSelect");
  const applyTemplateBtn = document.getElementById("applyTemplateBtn");
  if (applyTemplateBtn) {
    applyTemplateBtn.addEventListener("click", () => {
      const id = templateSelect.value;
      if (!id) {
        alert("請先選擇範本");
        return;
      }
      const template = templates.find((t) => String(t.id) === id);
      if (!template) return;
      if (!confirm(`套用範本「${template.name}」會覆蓋目前的品項與加料清單，確定嗎？`)) return;

      document.getElementById("storeInput").value = template.name;
      if (template.foodpanda_link) document.getElementById("foodpandaInput").value = template.foodpanda_link;
      if (template.notes) document.getElementById("notesInput").value = template.notes;

      clearRows(itemsRows);
      template.items.forEach((i) => makeRow(itemsRows, i.name, i.price));
      clearRows(toppingsRows);
      template.toppings.forEach((t) => makeRow(toppingsRows, t.name, t.price));
    });
  }

  // Delete template
  const deleteTemplateBtn = document.getElementById("deleteTemplateBtn");
  if (deleteTemplateBtn) {
    deleteTemplateBtn.addEventListener("click", () => {
      const id = templateSelect.value;
      if (!id) {
        alert("請先選擇範本");
        return;
      }
      const template = templates.find((t) => String(t.id) === id);
      if (!template) return;
      if (!confirm(`確定要刪除範本「${template.name}」嗎？此動作無法復原。`)) return;

      const form = document.createElement("form");
      form.method = "post";
      form.action = `/admin/templates/${id}/delete`;
      document.body.appendChild(form);
      form.submit();
    });
  }

  // Save current draft as a new template
  const saveTemplateBtn = document.getElementById("saveTemplateBtn");
  if (saveTemplateBtn) {
    saveTemplateBtn.addEventListener("click", () => {
      const name = prompt("範本名稱（建議用店家名稱）：", document.getElementById("storeInput").value);
      if (!name) return;

      const form = document.createElement("form");
      form.method = "post";
      form.action = "/admin/templates";
      form.style.display = "none";

      const fields = {
        name,
        foodpanda_link: document.getElementById("foodpandaInput").value,
        notes: document.getElementById("notesInput").value,
        items_json: JSON.stringify(collectRows(itemsRows)),
        toppings_json: JSON.stringify(collectRows(toppingsRows)),
      };
      Object.entries(fields).forEach(([key, value]) => {
        const input = document.createElement("input");
        input.type = "hidden";
        input.name = key;
        input.value = value;
        form.appendChild(input);
      });
      document.body.appendChild(form);
      form.submit();
    });
  }

  // Confirm before clearing the week's orders
  const resetForm = document.getElementById("resetForm");
  if (resetForm) {
    resetForm.addEventListener("submit", (e) => {
      if (!confirm("確定要清空本週訂單並開新一輪嗎？此動作無法復原。")) {
        e.preventDefault();
      }
    });
  }
})();
