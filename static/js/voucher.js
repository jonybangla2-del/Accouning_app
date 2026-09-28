// Voucher Posting Dynamic Row & Real-time Balance Calculation

document.addEventListener('DOMContentLoaded', () => {
    const tableBody = document.getElementById('voucher-lines-body');
    const addLineBtn = document.getElementById('add-line-btn');
    const voucherForm = document.getElementById('voucher-form');
    const totalDebitEl = document.getElementById('total-debit');
    const totalCreditEl = document.getElementById('total-credit');
    const balanceDiffEl = document.getElementById('balance-diff');
    const balanceStatusBadge = document.getElementById('balance-status-badge');
    const submitBtn = document.getElementById('submit-voucher-btn');

    // Voucher Type switcher (redirects to preset type when data-url is present)
    const voucherTypeSelect = document.getElementById('voucher_type');
    if (voucherTypeSelect && voucherTypeSelect.dataset.url) {
        voucherTypeSelect.addEventListener('change', function () {
            window.location.href = this.dataset.url + '?type=' + encodeURIComponent(this.value);
        });
    }

    if (!tableBody) return;

    // Track if user has manually modified Row 2 amounts
    let row2ManuallyEdited = false;

    // Default Row 2 to "Cash in Hand" if not already set
    function setDefaultCashInHandOnRow2() {
        const rows = tableBody.querySelectorAll('tr');
        if (rows.length >= 2) {
            const row2Select = rows[1].querySelector('.account-select');
            if (row2Select && (!row2Select.value || row2Select.selectedIndex <= 0)) {
                for (let i = 0; i < row2Select.options.length; i++) {
                    const optText = row2Select.options[i].text.toLowerCase();
                    if (optText.includes('cash in hand') || optText.includes('bs-001')) {
                        row2Select.selectedIndex = i;
                        break;
                    }
                }
            }
        }
    }
    setDefaultCashInHandOnRow2();

    // Recalculate totals and update balance status
    function calculateTotals() {
        let totalDebit = 0;
        let totalCredit = 0;

        const debitInputs = tableBody.querySelectorAll('.debit-input');
        const creditInputs = tableBody.querySelectorAll('.credit-input');

        debitInputs.forEach(input => {
            const val = parseFloat(input.value) || 0;
            totalDebit += val;
        });

        creditInputs.forEach(input => {
            const val = parseFloat(input.value) || 0;
            totalCredit += val;
        });

        totalDebit = Math.round(totalDebit * 100) / 100;
        totalCredit = Math.round(totalCredit * 100) / 100;

        const diff = Math.round(Math.abs(totalDebit - totalCredit) * 100) / 100;
        const isBalanced = diff < 0.005 && totalDebit > 0;

        if (totalDebitEl) totalDebitEl.textContent = '৳' + totalDebit.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        if (totalCreditEl) totalCreditEl.textContent = '৳' + totalCredit.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        if (balanceDiffEl) balanceDiffEl.textContent = '৳' + diff.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

        if (balanceStatusBadge) {
            if (isBalanced) {
                balanceStatusBadge.className = 'badge badge-balanced';
                balanceStatusBadge.innerHTML = '✓ Balanced';
                if (submitBtn) submitBtn.disabled = false;
            } else {
                balanceStatusBadge.className = 'badge badge-unbalanced';
                balanceStatusBadge.innerHTML = '⚠ Out of Balance';
                if (submitBtn) submitBtn.disabled = true;
            }
        }
    }

    // Attach event listeners to input fields
    function attachRowListeners(row) {
        const debitInput = row.querySelector('.debit-input');
        const creditInput = row.querySelector('.credit-input');
        const removeBtn = row.querySelector('.remove-line-btn');

        if (debitInput) {
            debitInput.addEventListener('input', () => {
                if (parseFloat(debitInput.value) > 0) {
                    creditInput.value = '';
                }

                // Requirement 2: Auto-populate 2nd line opposite side (Credit) when typing in 1st line Debit
                const allRows = tableBody.querySelectorAll('tr');
                if (row === allRows[0] && allRows.length >= 2 && !row2ManuallyEdited) {
                    const row2 = allRows[1];
                    const r2Debit = row2.querySelector('.debit-input');
                    const r2Credit = row2.querySelector('.credit-input');
                    const val = parseFloat(debitInput.value);

                    if (!isNaN(val) && val > 0) {
                        r2Credit.value = debitInput.value;
                        r2Debit.value = '';
                    } else if (!debitInput.value || val === 0) {
                        r2Credit.value = '';
                    }
                }

                calculateTotals();
            });
        }

        if (creditInput) {
            creditInput.addEventListener('input', () => {
                if (parseFloat(creditInput.value) > 0) {
                    debitInput.value = '';
                }

                // Requirement 2: Auto-populate 2nd line opposite side (Debit) when typing in 1st line Credit
                const allRows = tableBody.querySelectorAll('tr');
                if (row === allRows[0] && allRows.length >= 2 && !row2ManuallyEdited) {
                    const row2 = allRows[1];
                    const r2Debit = row2.querySelector('.debit-input');
                    const r2Credit = row2.querySelector('.credit-input');
                    const val = parseFloat(creditInput.value);

                    if (!isNaN(val) && val > 0) {
                        r2Debit.value = creditInput.value;
                        r2Credit.value = '';
                    } else if (!creditInput.value || val === 0) {
                        r2Debit.value = '';
                    }
                }

                calculateTotals();
            });
        }

        // Detect if user manually changes row 2 amounts
        const currentRows = tableBody.querySelectorAll('tr');
        if (currentRows.length >= 2 && row === currentRows[1]) {
            if (debitInput) {
                debitInput.addEventListener('keydown', (e) => {
                    if (e.key !== 'Tab' && e.key !== 'Shift' && e.key !== 'Enter') {
                        row2ManuallyEdited = true;
                    }
                });
            }
            if (creditInput) {
                creditInput.addEventListener('keydown', (e) => {
                    if (e.key !== 'Tab' && e.key !== 'Shift' && e.key !== 'Enter') {
                        row2ManuallyEdited = true;
                    }
                });
            }
        }

        if (removeBtn) {
            removeBtn.addEventListener('click', () => {
                const totalRows = tableBody.querySelectorAll('tr').length;
                if (totalRows <= 2) {
                    alert('A voucher must have at least 2 transaction lines for double-entry.');
                    return;
                }
                row.remove();
                calculateTotals();
            });
        }
    }

    // Initialize existing rows
    tableBody.querySelectorAll('tr').forEach(attachRowListeners);

    // Add row button
    if (addLineBtn) {
        addLineBtn.addEventListener('click', () => {
            const firstRow = tableBody.querySelector('tr');
            if (!firstRow) return;

            const newRow = firstRow.cloneNode(true);
            newRow.querySelectorAll('input').forEach(input => input.value = '');
            newRow.querySelector('select').selectedIndex = 0;

            tableBody.appendChild(newRow);
            attachRowListeners(newRow);

            // Auto-fill remaining unbalanced difference in the new row
            let totalDebit = 0, totalCredit = 0;
            tableBody.querySelectorAll('.debit-input').forEach(i => totalDebit += parseFloat(i.value) || 0);
            tableBody.querySelectorAll('.credit-input').forEach(i => totalCredit += parseFloat(i.value) || 0);
            const diff = Math.round((totalDebit - totalCredit) * 100) / 100;

            if (diff > 0) {
                newRow.querySelector('.credit-input').value = diff.toFixed(2);
            } else if (diff < 0) {
                newRow.querySelector('.debit-input').value = Math.abs(diff).toFixed(2);
            }

            calculateTotals();
        });
    }

    // Form submit validation
    if (voucherForm) {
        voucherForm.addEventListener('submit', (e) => {
            let totalDebit = 0;
            let totalCredit = 0;

            tableBody.querySelectorAll('.debit-input').forEach(i => totalDebit += parseFloat(i.value) || 0);
            tableBody.querySelectorAll('.credit-input').forEach(i => totalCredit += parseFloat(i.value) || 0);

            totalDebit = Math.round(totalDebit * 100) / 100;
            totalCredit = Math.round(totalCredit * 100) / 100;

            if (Math.abs(totalDebit - totalCredit) >= 0.01 || totalDebit <= 0) {
                e.preventDefault();
                alert('Voucher cannot be posted: Total Debits must equal Total Credits and be greater than 0.');
            }
        });
    }

    calculateTotals();
});
