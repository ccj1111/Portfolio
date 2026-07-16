(function () {
  const step1 = document.getElementById("step1");
  const step2 = document.getElementById("step2");
  if (!step1 || !step2) return; // menu not published / closed

  const NAME_KEY = "drinkOrderName";

  const state = {
    itemId: null,
    itemName: "",
    itemPrice: 0,
    qty: 1,
    toppings: [], // {id, name, price}
  };

  const qtyValueEl = document.getElementById("qtyValue");
  const selectedItemNameEl = document.getElementById("selectedItemName");
  const subtotalAmountEl = document.getElementById("subtotalAmount");
  const nameInput = document.getElementById("nameInput");
  const noteInput = document.getElementById("noteInput");
  const sugarSelect = document.getElementById("sugarSelect");
  const iceSelect = document.getElementById("iceSelect");

  const savedName = localStorage.getItem(NAME_KEY);
  if (savedName && nameInput) nameInput.value = savedName;

  function showToast(message) {
    const toast = document.getElementById("toast");
    toast.textContent = message;
    toast.classList.add("show");
    setTimeout(() => toast.classList.remove("show"), 2200);
  }

  function recomputeSubtotal() {
    const toppingTotal = state.toppings.reduce((sum, t) => sum + t.price, 0);
    const subtotal = (state.itemPrice + toppingTotal) * state.qty;
    subtotalAmountEl.textContent = "$" + subtotal;
    return subtotal;
  }

  function goToStep1() {
    step2.classList.remove("active");
    step1.classList.add("active");
  }

  function goToStep2() {
    step1.classList.remove("active");
    step2.classList.add("active");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function resetStep2Fields() {
    state.qty = 1;
    state.toppings = [];
    qtyValueEl.textContent = "1";
    if (sugarSelect) sugarSelect.value = "正常糖";
    if (iceSelect) iceSelect.value = "正常冰";
    if (noteInput) noteInput.value = "";
    document.querySelectorAll(".topping-tag.selected").forEach((el) => el.classList.remove("selected"));
  }

  document.querySelectorAll(".item-card").forEach((card) => {
    card.addEventListener("click", () => {
      state.itemId = card.dataset.id;
      state.itemName = card.dataset.name;
      state.itemPrice = parseInt(card.dataset.price, 10);
      resetStep2Fields();
      selectedItemNameEl.textContent = state.itemName + "（$" + state.itemPrice + "）";
      recomputeSubtotal();
      goToStep2();
    });
  });

  const backBtn = document.getElementById("backBtn");
  if (backBtn) backBtn.addEventListener("click", goToStep1);

  const qtyMinus = document.getElementById("qtyMinus");
  const qtyPlus = document.getElementById("qtyPlus");
  if (qtyMinus) {
    qtyMinus.addEventListener("click", () => {
      state.qty = Math.max(1, state.qty - 1);
      qtyValueEl.textContent = state.qty;
      recomputeSubtotal();
    });
  }
  if (qtyPlus) {
    qtyPlus.addEventListener("click", () => {
      state.qty += 1;
      qtyValueEl.textContent = state.qty;
      recomputeSubtotal();
    });
  }

  document.querySelectorAll(".topping-tag").forEach((tag) => {
    tag.addEventListener("click", () => {
      const id = tag.dataset.id;
      const idx = state.toppings.findIndex((t) => t.id === id);
      if (idx >= 0) {
        state.toppings.splice(idx, 1);
        tag.classList.remove("selected");
      } else {
        state.toppings.push({
          id,
          name: tag.dataset.name,
          price: parseInt(tag.dataset.price, 10),
        });
        tag.classList.add("selected");
      }
      recomputeSubtotal();
    });
  });

  const submitBtn = document.getElementById("submitBtn");
  if (submitBtn) {
    submitBtn.addEventListener("click", async () => {
      const personName = (nameInput.value || "").trim();
      if (!personName) {
        showToast("請填寫姓名");
        nameInput.focus();
        return;
      }
      if (!state.itemId) {
        showToast("請先選擇飲料");
        return;
      }

      submitBtn.disabled = true;
      try {
        const res = await fetch("/api/orders", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            person_name: personName,
            item_id: state.itemId,
            qty: state.qty,
            sugar: sugarSelect ? sugarSelect.value : "正常糖",
            ice: iceSelect ? iceSelect.value : "正常冰",
            topping_ids: state.toppings.map((t) => t.id),
            note: noteInput ? noteInput.value.trim() : "",
          }),
        });
        const data = await res.json();
        if (!res.ok || !data.ok) {
          showToast(data.error || "送出失敗，請再試一次");
          return;
        }
        localStorage.setItem(NAME_KEY, personName);
        showToast("送出成功！");
        state.itemId = null;
        goToStep1();
      } catch (err) {
        showToast("網路錯誤，請再試一次");
      } finally {
        submitBtn.disabled = false;
      }
    });
  }
})();
