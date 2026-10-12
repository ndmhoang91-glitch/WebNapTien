let currentMonth = "2026-10";
let currentTab = "otp"; // 'otp' | 'nap_ho' | 'chua_chay' | 'search'
let currentCustomerFilter = "ALL";
let currentOrderStatusFilter = "ALL";
let currentOrderOtpFilter = "ALL";
let currentOrders = [];
let currentPage = 1;
const pageSize = 100;
let selectedOrderIds = new Set();
let searchTimer = null;
let orderSearchTimer = null;
let dashboardData = null;

function formatVND(val) {
    const n = Number(val || 0);
    return n.toLocaleString("vi-VN") + " ₫";
}

function showToast(msg) {
    const container = document.getElementById("toastContainer");
    const el = document.createElement("div");
    el.className = "toast";
    el.textContent = msg;
    container.appendChild(el);
    setTimeout(() => {
        el.remove();
    }, 2200);
}

function copyText(text) {
    navigator.clipboard.writeText(text).then(() => {
        showToast(`Đã copy: ${text}`);
    });
}

async function initApp() {
    await loadMonths();
    await refreshAllData();
}

async function loadMonths() {
    const res = await fetch("/api/months");
    const data = await res.json();
    const select = document.getElementById("monthSelect");
    const copySelect = document.getElementById("copyFromMonthSelect");
    select.innerHTML = "";
    copySelect.innerHTML = `<option value="">-- Tạo tháng trống (Để tải file Excel mới lên) --</option>`;

    if (data.months && data.months.length > 0) {
        if (!data.months.some((m) => m.month_key === currentMonth)) {
            currentMonth = data.months[0].month_key;
        }
        data.months.forEach((m) => {
            const opt = document.createElement("option");
            opt.value = m.month_key;
            opt.textContent = `${m.label}`;
            if (m.month_key === currentMonth) opt.selected = true;
            select.appendChild(opt);

            const copyOpt = document.createElement("option");
            copyOpt.value = m.month_key;
            copyOpt.textContent = `Sao chép từ ${m.label}`;
            copySelect.appendChild(copyOpt);
        });
    }

    select.onchange = async (e) => {
        currentMonth = e.target.value;
        currentPage = 1;
        selectedOrderIds.clear();
        await refreshAllData();
    };
}

async function refreshAllData() {
    await loadDashboard();
    if (currentTab === "search") {
        await loadCustomerSummary();
    } else {
        renderOrderFilterButtons();
        updateBulkBarVisibility();
        await loadOrders(currentPage);
    }
}

async function loadDashboard() {
    const res = await fetch(`/api/dashboard?month_key=${encodeURIComponent(currentMonth)}`);
    const d = await res.json();
    dashboardData = d;

    document.getElementById("kpiPaidAmount").textContent = formatVND(d.paid_amount);
    document.getElementById("kpiPaidOrders").textContent = d.paid_orders.toLocaleString("vi-VN");
    document.getElementById("kpiUnpaidAmount").textContent = formatVND(d.unpaid_amount);
    document.getElementById("kpiUnpaidOrders").textContent = d.unpaid_orders.toLocaleString("vi-VN");

    document.getElementById("kpiNeedOtpCount").textContent = `${d.otp.need_otp_orders} số`;
    document.getElementById("kpiNeedOtpAmount").textContent = formatVND(d.otp.need_otp_amount);
    document.getElementById("kpiHasOtpCount").textContent = `${d.otp.has_otp_orders} số`;
    document.getElementById("kpiHasOtpAmount").textContent = formatVND(d.otp.has_otp_amount);

    document.getElementById("countTabOtp").textContent = d.otp.total_orders;
    document.getElementById("countTabNapHo").textContent = d.nap_ho.total_orders;
    document.getElementById("countTabChuaChay").textContent = d.chua_chay.total_orders;
}

function switchTab(tab) {
    currentTab = tab;
    document.querySelectorAll(".sheet-tab").forEach((btn) => {
        btn.classList.toggle("active", btn.dataset.tab === tab);
    });

    if (tab === "search") {
        document.getElementById("panel-orders").classList.remove("active");
        document.getElementById("panel-search").classList.add("active");
        loadCustomerSummary();
    } else {
        document.getElementById("panel-search").classList.remove("active");
        document.getElementById("panel-orders").classList.add("active");
        currentOrderStatusFilter = "ALL";
        currentOrderOtpFilter = "ALL";
        currentPage = 1;
        selectedOrderIds.clear();
        renderOrderFilterButtons();
        updateBulkBarVisibility();
        loadOrders(1);
    }
}

function quickFilterOtp(otpStatus) {
    currentTab = "otp";
    document.querySelectorAll(".sheet-tab").forEach((btn) => {
        btn.classList.toggle("active", btn.dataset.tab === "otp");
    });
    document.getElementById("panel-search").classList.remove("active");
    document.getElementById("panel-orders").classList.add("active");
    currentOrderStatusFilter = "PENDING";
    currentOrderOtpFilter = otpStatus;
    currentPage = 1;
    selectedOrderIds.clear();
    renderOrderFilterButtons();
    updateBulkBarVisibility();
    loadOrders(1);
}

function renderOrderFilterButtons() {
    const container = document.getElementById("orderFilterButtons");
    const d = dashboardData;
    if (currentTab === "otp") {
        const allCnt = d ? d.otp.total_orders : 0;
        const needCnt = d ? d.otp.need_otp_orders : 0;
        const hasCnt = d ? d.otp.has_otp_orders : 0;
        const paidCnt = d ? d.otp.paid_orders : 0;
        container.innerHTML = `
            <button class="filter-btn ${currentOrderStatusFilter === 'ALL' && currentOrderOtpFilter === 'ALL' ? 'active' : ''}"
                onclick="applyOrderSubFilter('ALL', 'ALL')">Tất cả (${allCnt})</button>
            <button class="filter-btn ${currentOrderOtpFilter === 'NEED_OTP' ? 'active' : ''}"
                onclick="applyOrderSubFilter('PENDING', 'NEED_OTP')">🔴 Lấy OTP (${needCnt})</button>
            <button class="filter-btn ${currentOrderOtpFilter === 'HAS_OTP' ? 'active' : ''}"
                onclick="applyOrderSubFilter('PENDING', 'HAS_OTP')">🔵 Đã có OTP (${hasCnt})</button>
            <button class="filter-btn ${currentOrderStatusFilter === 'COMPLETED' ? 'active' : ''}"
                onclick="applyOrderSubFilter('COMPLETED', 'ALL')">🟢 Đã nạp (${paidCnt})</button>
        `;
    } else if (currentTab === "nap_ho") {
        const allCnt = d ? d.nap_ho.total_orders : 0;
        const unCnt = d ? d.nap_ho.unpaid_orders : 0;
        const paidCnt = d ? d.nap_ho.paid_orders : 0;
        container.innerHTML = `
            <button class="filter-btn ${currentOrderStatusFilter === 'ALL' ? 'active' : ''}"
                onclick="applyOrderSubFilter('ALL', 'ALL')">Tất cả (${allCnt})</button>
            <button class="filter-btn ${currentOrderStatusFilter === 'PENDING' ? 'active' : ''}"
                onclick="applyOrderSubFilter('PENDING', 'ALL')">🟠 Chưa nạp (${unCnt})</button>
            <button class="filter-btn ${currentOrderStatusFilter === 'COMPLETED' ? 'active' : ''}"
                onclick="applyOrderSubFilter('COMPLETED', 'ALL')">🟢 Đã nạp (${paidCnt})</button>
        `;
    } else {
        const allCnt = d ? d.chua_chay.total_orders : 0;
        const holdCnt = d ? d.chua_chay.holding_orders : 0;
        const doneCnt = d ? d.chua_chay.scheduled_orders : 0;
        container.innerHTML = `
            <button class="filter-btn ${currentOrderStatusFilter === 'ALL' ? 'active' : ''}"
                onclick="applyOrderSubFilter('ALL', 'ALL')">Tất cả (${allCnt})</button>
            <button class="filter-btn ${currentOrderStatusFilter === 'PENDING' ? 'active' : ''}"
                onclick="applyOrderSubFilter('PENDING', 'ALL')">Chưa cho lên đơn (${holdCnt})</button>
            <button class="filter-btn ${currentOrderStatusFilter === 'COMPLETED' ? 'active' : ''}"
                onclick="applyOrderSubFilter('COMPLETED', 'ALL')">Đã lên đơn (${doneCnt})</button>
        `;
    }
}

function applyOrderSubFilter(status, otpStatus) {
    currentOrderStatusFilter = status;
    currentOrderOtpFilter = otpStatus;
    currentPage = 1;
    renderOrderFilterButtons();
    loadOrders(1);
}

/* ==================== ORDERS TABLE ==================== */
function debounceLoadOrders() {
    clearTimeout(orderSearchTimer);
    orderSearchTimer = setTimeout(() => {
        loadOrders(1);
    }, 200);
}

async function loadOrders(page = 1) {
    currentPage = page;
    const orderType = currentTab === "otp" ? "OTP" : currentTab === "nap_ho" ? "NAP_HO" : "CHUA_CHAY";
    const q = document.getElementById("orderSearchInput").value.trim();

    const params = new URLSearchParams({
        month_key: currentMonth,
        order_type: orderType,
        status: currentOrderStatusFilter,
        otp_status: currentOrderOtpFilter,
        q: q,
        page: page,
        page_size: pageSize,
    });

    const res = await fetch(`/api/orders?${params.toString()}`);
    const data = await res.json();
    currentOrders = data.orders || [];

    document.getElementById("filterTotalCount").textContent = data.total.toLocaleString("vi-VN");
    document.getElementById("filterSumPaid").textContent = formatVND(data.sum_paid);
    document.getElementById("filterSumUnpaid").textContent = formatVND(data.sum_unpaid);

    renderOrdersTable();
    renderPagination(data.total, data.page, data.page_size);
}

function renderOrdersTable() {
    const tbody = document.getElementById("ordersTableBody");
    if (!currentOrders || currentOrders.length === 0) {
        tbody.innerHTML = `<tr><td colspan="8" class="text-center" style="padding:20px;color:#6b7280;">Không có dòng dữ liệu nào.</td></tr>`;
        updateSelectedCountUI();
        return;
    }

    tbody.innerHTML = currentOrders
        .map((o, idx) => {
            const stt = (currentPage - 1) * pageSize + idx + 1;
            const isChecked = selectedOrderIds.has(o.id) ? "checked" : "";

            let rowClass = "";
            if (o.status === "COMPLETED") {
                rowClass = "row-completed";
            } else if (o.order_type === "OTP" && o.otp_status === "HAS_OTP") {
                rowClass = "row-has-otp";
            } else if (o.order_type === "OTP" && o.otp_status === "NEED_OTP") {
                rowClass = "row-need-otp";
            }

            let statusButtons = "";
            if (o.order_type === "OTP") {
                statusButtons = `
                    <div class="state-switch-group">
                        <button class="state-btn ${o.status !== 'COMPLETED' && o.otp_status === 'NEED_OTP' ? 'active-need' : ''}"
                            onclick="updateOrderState(${o.id}, {status:'PENDING', otp_status:'NEED_OTP'})">Lấy OTP</button>
                        <button class="state-btn ${o.status !== 'COMPLETED' && o.otp_status === 'HAS_OTP' ? 'active-has' : ''}"
                            onclick="updateOrderState(${o.id}, {status:'PENDING', otp_status:'HAS_OTP'})">Đã có OTP</button>
                        <button class="state-btn ${o.status === 'COMPLETED' ? 'active-done' : ''}"
                            onclick="updateOrderState(${o.id}, {status:'COMPLETED', otp_status:'DONE_OTP'})">Đã nạp</button>
                    </div>
                `;
            } else if (o.order_type === "NAP_HO") {
                statusButtons = `
                    <div class="state-switch-group">
                        <button class="state-btn ${o.status !== 'COMPLETED' ? 'active-pending' : ''}"
                            onclick="updateOrderState(${o.id}, {status:'PENDING'})">Chưa nạp</button>
                        <button class="state-btn ${o.status === 'COMPLETED' ? 'active-done' : ''}"
                            onclick="updateOrderState(${o.id}, {status:'COMPLETED'})">Đã nạp</button>
                    </div>
                `;
            } else {
                statusButtons = `
                    <div class="state-switch-group">
                        <button class="state-btn ${o.status !== 'COMPLETED' ? 'active-pending' : ''}"
                            onclick="updateOrderState(${o.id}, {status:'PENDING'})">Chưa cho lên đơn</button>
                        <button class="state-btn ${o.status === 'COMPLETED' ? 'active-done' : ''}"
                            onclick="updateOrderState(${o.id}, {status:'COMPLETED'})">Đã lên đơn</button>
                    </div>
                `;
            }

            return `
            <tr class="${rowClass}">
                <td class="text-center">
                    <input type="checkbox" class="order-checkbox" value="${o.id}" ${isChecked} onchange="onSelectOrderCheckbox(this)">
                </td>
                <td>${stt}</td>
                <td>
                    <strong style="font-family:monospace;cursor:pointer;" onclick="copyText('${o.phone}')" title="Bấm để copy SĐT">${o.phone}</strong>
                </td>
                <td>
                    <a href="javascript:void(0)" onclick="openCustomerDetail('${o.phone}')" style="color:#111827;font-weight:600;text-decoration:none;">
                        ${escapeHtml(o.customer_name)}
                    </a>
                </td>
                <td class="text-center">${statusButtons}</td>
                <td class="text-right">
                    <input type="text" class="cell-input-amount" value="${o.amount.toLocaleString('vi-VN')}"
                        onblur="onInlineAmountChange(${o.id}, this)"
                        onkeydown="if(event.key==='Enter'){this.blur();}"
                        title="Bấm để sửa số tiền">
                </td>
                <td>
                    <input type="text" class="cell-input-note" placeholder="..." value="${escapeHtml(o.note || '')}"
                        onblur="updateOrderField(${o.id}, 'note', this.value)"
                        onkeydown="if(event.key==='Enter'){this.blur();}">
                </td>
                <td class="text-center">
                    <button class="btn btn-xs btn-outline-danger" onclick="deleteSingleOrder(${o.id})">✕</button>
                </td>
            </tr>
        `;
        })
        .join("");

    updateSelectedCountUI();
}

function renderPagination(total, page, size) {
    const totalPages = Math.max(1, Math.ceil(total / size));
    document.getElementById("paginationInfo").textContent = `Trang ${page} / ${totalPages} (Tổng ${total} dòng)`;
    const controls = document.getElementById("paginationControls");
    let html = "";
    if (page > 1) {
        html += `<button class="btn btn-xs btn-outline" onclick="loadOrders(${page - 1})">« Trước</button>`;
    }
    const startP = Math.max(1, page - 2);
    const endP = Math.min(totalPages, page + 2);
    for (let p = startP; p <= endP; p++) {
        html += `<button class="btn btn-xs ${p === page ? 'btn-primary' : 'btn-outline'}" onclick="loadOrders(${p})">${p}</button>`;
    }
    if (page < totalPages) {
        html += `<button class="btn btn-xs btn-outline" onclick="loadOrders(${page + 1})">Sau »</button>`;
    }
    controls.innerHTML = html;
}

async function updateOrderState(orderId, payload) {
    const res = await fetch(`/api/orders/${orderId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (data.success) {
        await loadDashboard();
        renderOrderFilterButtons();
        await loadOrders(currentPage);
    }
}

async function onInlineAmountChange(orderId, inputEl) {
    const raw = inputEl.value.replace(/[^\d]/g, "");
    const newAmt = parseInt(raw || "0", 10);
    const order = currentOrders.find((o) => o.id === orderId);
    if (!order || order.amount === newAmt) {
        inputEl.value = (order ? order.amount : newAmt).toLocaleString("vi-VN");
        return;
    }
    const res = await fetch(`/api/orders/${orderId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ amount: newAmt }),
    });
    const data = await res.json();
    if (data.success) {
        order.amount = data.amount;
        inputEl.value = data.amount.toLocaleString("vi-VN");
        showToast(`Đã lưu số tiền: ${formatVND(data.amount)}`);
        await loadDashboard();
        await loadOrders(currentPage);
    }
}

async function updateOrderField(orderId, field, value) {
    const order = currentOrders.find((o) => o.id === orderId);
    if (order && (order[field] || "") === value.trim()) return;
    await fetch(`/api/orders/${orderId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ [field]: value.trim() }),
    });
    if (order) order[field] = value.trim();
    showToast("Đã lưu ghi chú");
}

async function deleteSingleOrder(orderId) {
    if (!confirm("Xóa dòng đơn này?")) return;
    await fetch("/api/orders/bulk", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ids: [orderId], action: "delete" }),
    });
    selectedOrderIds.delete(orderId);
    await refreshAllData();
}

function onSelectOrderCheckbox(cb) {
    const id = parseInt(cb.value, 10);
    if (cb.checked) selectedOrderIds.add(id);
    else selectedOrderIds.delete(id);
    updateSelectedCountUI();
}

function toggleSelectAll(masterCb) {
    currentOrders.forEach((o) => {
        if (masterCb.checked) selectedOrderIds.add(o.id);
        else selectedOrderIds.delete(o.id);
    });
    document.querySelectorAll(".order-checkbox").forEach((cb) => {
        cb.checked = masterCb.checked;
    });
    updateSelectedCountUI();
}

function updateBulkBarVisibility() {
    const isOtp = currentTab === "otp";
    const isChuaChay = currentTab === "chua_chay";
    document.getElementById("btnBulkHasOtp").style.display = isOtp ? "inline-block" : "none";
    document.getElementById("btnBulkNeedOtp").style.display = isOtp ? "inline-block" : "none";
    document.getElementById("btnBulkMoveOtp").style.display = isChuaChay ? "inline-block" : "none";
    document.getElementById("btnBulkMoveNapHo").style.display = isChuaChay ? "inline-block" : "none";
}

function updateSelectedCountUI() {
    document.getElementById("selectedCount").textContent = selectedOrderIds.size;
    document.getElementById("bulkActionBar").style.display = selectedOrderIds.size > 0 ? "flex" : "none";
}

async function bulkAction(action) {
    if (selectedOrderIds.size === 0) return;
    if (action === "delete" && !confirm(`Xóa ${selectedOrderIds.size} dòng đã chọn?`)) return;

    const res = await fetch("/api/orders/bulk", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ids: Array.from(selectedOrderIds), action }),
    });
    const data = await res.json();
    if (data.success) {
        selectedOrderIds.clear();
        document.getElementById("selectAllCheckbox").checked = false;
        await refreshAllData();
    }
}

/* ==================== TAB TRA CỨU KHÁCH HÀNG ==================== */
function debounceLoadCustomerSummary() {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
        loadCustomerSummary();
    }, 200);
}

function setCustomerFilter(filterMode) {
    currentCustomerFilter = filterMode;
    document.querySelectorAll("#custFilterPills .filter-btn").forEach((p) => {
        p.classList.toggle("active", p.dataset.filter === filterMode);
    });
    loadCustomerSummary();
}

async function loadCustomerSummary() {
    const q = document.getElementById("custSearchInput").value.trim();
    const url = `/api/customers/summary?month_key=${encodeURIComponent(currentMonth)}&q=${encodeURIComponent(q)}&filter=${encodeURIComponent(currentCustomerFilter)}`;
    const res = await fetch(url);
    const data = await res.json();
    const tbody = document.getElementById("customerSummaryBody");

    if (!data.customers || data.customers.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="text-center" style="padding:20px;color:#6b7280;">Không tìm thấy khách hàng nào.</td></tr>`;
        return;
    }

    tbody.innerHTML = data.customers
        .slice(0, 300)
        .map((c, idx) => {
            const badges = [];
            if (c.need_otp_count > 0) {
                badges.push(`<span class="badge badge-danger">🔴 Cần lấy OTP: ${c.need_otp_count} (${formatVND(c.need_otp_amount)})</span>`);
            }
            if (c.has_otp_count > 0) {
                badges.push(`<span class="badge badge-info">🔵 Đã có OTP: ${c.has_otp_count} (${formatVND(c.has_otp_amount)})</span>`);
            }
            if (c.napho_unpaid_count > 0) {
                badges.push(`<span class="badge badge-warning">Nạp hộ chưa nạp: ${c.napho_unpaid_count} (${formatVND(c.napho_unpaid_amount)})</span>`);
            }
            if (badges.length === 0 && c.total_paid > 0) {
                badges.push(`<span class="badge badge-success">Đã nạp xong</span>`);
            }

            return `
            <tr>
                <td>${idx + 1}</td>
                <td><strong style="font-family:monospace;">${c.phone}</strong></td>
                <td>
                    <a href="javascript:void(0)" onclick="openCustomerDetail('${c.phone}')" style="font-weight:700;color:#1e3a8a;text-decoration:none;">
                        ${escapeHtml(c.customer_name)}
                    </a>
                </td>
                <td class="text-right"><strong class="text-success">${formatVND(c.total_paid)}</strong></td>
                <td class="text-right">${c.otp_paid > 0 ? formatVND(c.otp_paid) : "-"}</td>
                <td class="text-right">${c.napho_paid > 0 ? formatVND(c.napho_paid) : "-"}</td>
                <td class="text-right">${c.total_unpaid > 0 ? `<span class="text-danger">${formatVND(c.total_unpaid)}</span>` : "-"}</td>
                <td><div style="display:flex;gap:4px;flex-wrap:wrap;">${badges.join("")}</div></td>
                <td class="text-center">
                    <button class="btn btn-xs btn-primary" onclick="openCustomerDetail('${c.phone}')">Xem</button>
                </td>
            </tr>
        `;
        })
        .join("");
}

async function openCustomerDetail(phone) {
    const res = await fetch(`/api/customers/${encodeURIComponent(phone)}/detail?month_key=${encodeURIComponent(currentMonth)}`);
    const data = await res.json();
    const cust = data.customer;
    document.getElementById("custModalTitle").textContent = `${cust.name} - ${cust.phone}`;

    const totalPaid = data.month_orders
        .filter((o) => (o.order_type === "OTP" || o.order_type === "NAP_HO") && o.status === "COMPLETED")
        .reduce((acc, o) => acc + o.amount, 0);
    const totalUnpaid = data.month_orders
        .filter((o) => (o.order_type === "OTP" || o.order_type === "NAP_HO") && o.status !== "COMPLETED")
        .reduce((acc, o) => acc + o.amount, 0);

    const ordersRowsHtml = data.month_orders
        .map((o, i) => {
            const typeLabel = o.order_type === "OTP" ? "Đơn OTP" : o.order_type === "NAP_HO" ? "Đơn Nạp hộ" : "Đơn Chưa Cho Chạy";
            let statusControl = "";
            if (o.order_type === "OTP") {
                statusControl = `
                    <div class="state-switch-group">
                        <button class="state-btn ${o.status !== 'COMPLETED' && o.otp_status === 'NEED_OTP' ? 'active-need' : ''}"
                            onclick="quickUpdateOrderInModal(${o.id}, {status:'PENDING', otp_status:'NEED_OTP'}, '${cust.phone}')">Lấy OTP</button>
                        <button class="state-btn ${o.status !== 'COMPLETED' && o.otp_status === 'HAS_OTP' ? 'active-has' : ''}"
                            onclick="quickUpdateOrderInModal(${o.id}, {status:'PENDING', otp_status:'HAS_OTP'}, '${cust.phone}')">Đã có OTP</button>
                        <button class="state-btn ${o.status === 'COMPLETED' ? 'active-done' : ''}"
                            onclick="quickUpdateOrderInModal(${o.id}, {status:'COMPLETED', otp_status:'DONE_OTP'}, '${cust.phone}')">Đã nạp</button>
                    </div>
                `;
            } else {
                statusControl = `
                    <div class="state-switch-group">
                        <button class="state-btn ${o.status !== 'COMPLETED' ? 'active-pending' : ''}"
                            onclick="quickUpdateOrderInModal(${o.id}, {status:'PENDING'}, '${cust.phone}')">Chưa nạp</button>
                        <button class="state-btn ${o.status === 'COMPLETED' ? 'active-done' : ''}"
                            onclick="quickUpdateOrderInModal(${o.id}, {status:'COMPLETED'}, '${cust.phone}')">Đã nạp</button>
                    </div>
                `;
            }

            return `
            <tr>
                <td>${i + 1}</td>
                <td>${typeLabel}</td>
                <td class="text-right"><strong>${formatVND(o.amount)}</strong></td>
                <td class="text-center">${statusControl}</td>
                <td>${escapeHtml(o.note || "")}</td>
            </tr>
        `;
        })
        .join("");

    const historyHtml = (data.history || [])
        .map(
            (h) => `
        <tr>
            <td><strong>${h.label || h.month_key}</strong></td>
            <td class="text-center">${h.order_count} đơn</td>
            <td class="text-right text-success">${formatVND(h.paid_amount)}</td>
            <td class="text-right text-danger">${formatVND(h.unpaid_amount)}</td>
        </tr>
    `
        )
        .join("");

    document.getElementById("custModalBody").innerHTML = `
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px;">
            <div style="background:#f0fdf4;border:1px solid #bbf7d0;padding:10px 14px;border-radius:6px;">
                <div style="font-size:12px;color:#166534;">ĐÃ NẠP TRONG THÁNG (${currentMonth})</div>
                <div style="font-size:20px;font-weight:700;color:#15803d;">${formatVND(totalPaid)}</div>
            </div>
            <div style="background:#fff1f2;border:1px solid #fecdd3;padding:10px 14px;border-radius:6px;">
                <div style="font-size:12px;color:#9f1239;">CHƯA NẠP (${currentMonth})</div>
                <div style="font-size:20px;font-weight:700;color:#be123c;">${formatVND(totalUnpaid)}</div>
            </div>
        </div>
        <h4 style="margin-bottom:6px;">Các đơn trong tháng ${currentMonth}</h4>
        <div class="table-wrap" style="margin-bottom:16px;">
            <table class="excel-table">
                <thead>
                    <tr>
                        <th>STT</th>
                        <th>Sheet / Loại</th>
                        <th class="text-right">Số Tiền</th>
                        <th class="text-center">Trạng Thái</th>
                        <th>Ghi Chú</th>
                    </tr>
                </thead>
                <tbody>
                    ${ordersRowsHtml || '<tr><td colspan="5" class="text-center">Không có đơn</td></tr>'}
                </tbody>
            </table>
        </div>
        <h4 style="margin-bottom:6px;">Lịch sử qua các tháng</h4>
        <div class="table-wrap">
            <table class="excel-table">
                <thead>
                    <tr>
                        <th>Tháng</th>
                        <th class="text-center">Số đơn</th>
                        <th class="text-right">Đã nạp</th>
                        <th class="text-right">Chưa nạp</th>
                    </tr>
                </thead>
                <tbody>
                    ${historyHtml || '<tr><td colspan="4" class="text-center">Chưa có lịch sử</td></tr>'}
                </tbody>
            </table>
        </div>
    `;

    openModal("customerDetailModal");
}

async function quickUpdateOrderInModal(orderId, payload, phone) {
    await fetch(`/api/orders/${orderId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
    });
    await openCustomerDetail(phone);
    await refreshAllData();
}

/* ==================== MODALS ==================== */
function openModal(id) {
    document.getElementById(id).style.display = "flex";
}

function closeModal(id) {
    document.getElementById(id).style.display = "none";
}

function openImportModal() {
    document.getElementById("importMonthInput").value = currentMonth;
    openModal("importExcelModal");
}

async function submitImportExcel(e) {
    e.preventDefault();
    const fileInput = document.getElementById("importFileInput");
    if (!fileInput.files.length) return;

    const btn = document.getElementById("btnSubmitImport");
    btn.disabled = true;
    btn.textContent = "Đang nhập...";

    const formData = new FormData();
    formData.append("file", fileInput.files[0]);
    formData.append("month_key", document.getElementById("importMonthInput").value);
    formData.append("mode", document.getElementById("importModeSelect").value);
    formData.append("default_order_type", document.getElementById("importDefaultTypeSelect").value);

    try {
        const res = await fetch("/api/import-excel", {
            method: "POST",
            body: formData,
        });
        const data = await res.json();
        if (data.success) {
            currentMonth = document.getElementById("importMonthInput").value;
            closeModal("importExcelModal");
            showToast(`Đã nhập ${data.imported_orders} dòng đơn!`);
            await loadMonths();
            await refreshAllData();
        } else {
            alert(data.error || "Lỗi nhập file Excel");
        }
    } finally {
        btn.disabled = false;
        btn.textContent = "Nhập Dữ Liệu";
    }
}

function exportCurrentMonthExcel() {
    window.location.href = `/api/export-excel?month_key=${encodeURIComponent(currentMonth)}`;
}

function openNewMonthModal() {
    const now = new Date();
    const nextM = String(now.getMonth() + 1).padStart(2, "0");
    document.getElementById("newMonthKeyInput").value = `${now.getFullYear()}-${nextM}`;
    openModal("newMonthModal");
}

async function submitNewMonth(e) {
    e.preventDefault();
    const monthKey = document.getElementById("newMonthKeyInput").value;
    const copyFrom = document.getElementById("copyFromMonthSelect").value;
    const resetStatus = document.getElementById("resetStatusCheckbox").checked;

    const res = await fetch("/api/months", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ month_key: monthKey, copy_from: copyFrom, reset_status: resetStatus }),
    });
    const data = await res.json();
    if (data.success) {
        currentMonth = data.month_key;
        closeModal("newMonthModal");
        showToast(`Đã tạo ${data.label}`);
        await loadMonths();
        await refreshAllData();
    } else {
        alert(data.error || "Không thể tạo tháng");
    }
}

function openAddOrderModal() {
    document.getElementById("addOrderPhone").value = "";
    document.getElementById("addOrderName").value = "";
    document.getElementById("addOrderAmount").value = "";
    document.getElementById("addOrderNote").value = "";
    if (currentTab === "nap_ho") document.getElementById("addOrderType").value = "NAP_HO";
    else if (currentTab === "chua_chay") document.getElementById("addOrderType").value = "CHUA_CHAY";
    else document.getElementById("addOrderType").value = "OTP";
    toggleAddOrderOtpField();
    openModal("addOrderModal");
}

function toggleAddOrderOtpField() {
    const t = document.getElementById("addOrderType").value;
    document.getElementById("addOrderOtpGroup").style.display = t === "OTP" ? "block" : "none";
}

async function autoFillCustomerName() {
    const phone = document.getElementById("addOrderPhone").value.trim();
    if (!phone) return;
    const res = await fetch(`/api/customers/${encodeURIComponent(phone)}/detail?month_key=${encodeURIComponent(currentMonth)}`);
    const data = await res.json();
    if (data.customer && data.customer.name && data.customer.name !== "Khách hàng") {
        const nameInput = document.getElementById("addOrderName");
        if (!nameInput.value.trim()) {
            nameInput.value = data.customer.name;
        }
    }
}

async function submitAddOrder(e) {
    e.preventDefault();
    const orderType = document.getElementById("addOrderType").value;
    const otpStatus = document.getElementById("addOrderOtpStatus").value;
    const status = orderType === "OTP" && otpStatus === "DONE_OTP" ? "COMPLETED" : "PENDING";

    const payload = {
        month_key: currentMonth,
        phone: document.getElementById("addOrderPhone").value,
        customer_name: document.getElementById("addOrderName").value,
        amount: document.getElementById("addOrderAmount").value,
        order_type: orderType,
        status: status,
        otp_status: orderType === "OTP" ? otpStatus : "NONE",
        note: document.getElementById("addOrderNote").value,
    };

    const res = await fetch("/api/orders", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (data.success) {
        closeModal("addOrderModal");
        showToast("Đã thêm đơn!");
        await refreshAllData();
    } else {
        alert(data.error || "Lỗi thêm đơn");
    }
}

function openAddCustomerModal() {
    document.getElementById("custFormPhone").value = "";
    document.getElementById("custFormName").value = "";
    openModal("addCustomerModal");
}

async function submitUpsertCustomer(e) {
    e.preventDefault();
    const res = await fetch("/api/customers", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            phone: document.getElementById("custFormPhone").value,
            name: document.getElementById("custFormName").value,
        }),
    });
    const data = await res.json();
    if (data.success) {
        closeModal("addCustomerModal");
        showToast(`Đã lưu KH: ${data.name}`);
        await refreshAllData();
    } else {
        alert(data.error || "Lỗi lưu khách hàng");
    }
}

function escapeHtml(str) {
    return String(str || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

/* ==================== LIÊN KẾT TỰ ĐỘNG BC26.PRO ==================== */
let lastKnownBc26Version = 0;

function initBc26Bookmarklet() {
    const serverOrigin = window.location.origin;
    const bookmarkletCode = `javascript:(function(){if(window._bc26SyncTimer){alert('Đang chạy đồng bộ tự động với Web Quản Lý rồi!');return;}function scan(){var items=[];document.querySelectorAll('table tr').forEach(function(tr){var t=tr.innerText||'';var m=t.match(/\\b(0[35789]\\d{8})\\b/);if(!m)return;var hasOtpInput=tr.querySelector('input[placeholder*="OTP" i]')!==null||/nhập\\s*OTP/i.test(t)||/OK/.test(t)&&/OTP/.test(t);var st='NEED_OTP',note='Chưa nhập OTP (bc26)';if(/Hoàn thành|Thành công/i.test(t)){st='COMPLETED';note='Đã nạp trên bc26';}else if(!hasOtpInput&&/Chờ xử lý|Đang xử lý|Đã nhận/i.test(t)){st='HAS_OTP';note='Đã có OTP - Chờ xử lý';}else if(!hasOtpInput){return;}items.push({phone:m[1],state:st,status_text:note});});if(items.length>0){fetch('${serverOrigin}/api/bc26/sync',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({items:items})}).then(function(r){return r.json();}).then(function(d){var b=document.getElementById('_bc26_badge');if(!b){b=document.createElement('div');b.id='_bc26_badge';b.style.cssText='position:fixed;bottom:10px;right:10px;z-index:999999;background:#16a34a;color:#fff;padding:8px 14px;border-radius:6px;font-weight:bold;font-size:13px;box-shadow:0 2px 8px rgba(0,0,0,0.3);';document.body.appendChild(b);}b.textContent='🟢 Đang đồng bộ tự động '+items.length+' số sang Web Quản Lý (Đã có OTP: '+d.has_otp_count+' | Cần OTP: '+d.need_otp_count+')';});}}scan();window._bc26SyncTimer=setInterval(scan,3000);})();`;
    const link = document.getElementById("bc26BookmarkletLink");
    if (link) {
        link.href = bookmarkletCode;
    }
}

function openBc26Modal() {
    initBc26Bookmarklet();
    openModal("bc26Modal");
}

async function pollBc26SyncStatus() {
    try {
        const res = await fetch("/api/bc26/status");
        const data = await res.json();
        if (data && data.sync_version && data.sync_version !== lastKnownBc26Version) {
            const isFirstLoad = lastKnownBc26Version === 0;
            lastKnownBc26Version = data.sync_version;
            const btn = document.getElementById("btnBc26SyncStatus");
            if (btn && data.last_sync_at) {
                btn.className = "btn btn-success";
                btn.innerHTML = `🟢 Đang đồng bộ bc26 (${data.last_sync_at.split(" ")[0]})`;
            }
            const modalTxt = document.getElementById("bc26ModalStatusText");
            if (modalTxt && data.last_sync_at) {
                modalTxt.textContent = `Đã đồng bộ ${data.total_read} số lúc ${data.last_sync_at} (Đã có OTP: ${data.has_otp_count} | Cần lấy OTP: ${data.need_otp_count} | Đã nạp: ${data.completed_count})`;
            }
            if (!isFirstLoad) {
                await refreshAllData();
            }
        }
    } catch (e) {}
}

async function parseAndSyncPastedBc26() {
    const raw = document.getElementById("bc26PasteArea").value || "";
    const lines = raw.split(/\r?\n/);
    const items = [];
    for (const line of lines) {
        const m = line.match(/\b(0[35789]\d{8})\b/);
        if (!m) continue;
        const phone = m[1];
        let state = "NEED_OTP";
        let status_text = "Chưa nhập OTP (bc26)";
        if (/Hoàn thành|Thành công/i.test(line)) {
            state = "COMPLETED";
            status_text = "Đã nạp trên bc26";
        } else if (/Chờ xử lý|Đang xử lý|Đã nhận/i.test(line) && !/nhập\s*OTP/i.test(line)) {
            state = "HAS_OTP";
            status_text = "Đã có OTP - Chờ xử lý";
        }
        items.push({ phone, state, status_text });
    }
    if (items.length === 0) {
        alert("Không tìm thấy số điện thoại nào trong nội dung vừa dán!");
        return;
    }
    const res = await fetch("/api/bc26/sync", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ month_key: currentMonth, items }),
    });
    const data = await res.json();
    if (data.success) {
        document.getElementById("bc26PasteArea").value = "";
        closeModal("bc26Modal");
        showToast(`✅ Đã đồng bộ ${items.length} số từ bc26.pro (Cập nhật ${data.updated_count} đơn)!`);
        await refreshAllData();
    }
}

window.addEventListener("DOMContentLoaded", () => {
    initApp();
    initBc26Bookmarklet();
    setInterval(pollBc26SyncStatus, 3000);
});
