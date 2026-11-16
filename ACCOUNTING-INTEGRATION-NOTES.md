# Integrated accounting workflows — implementation notes

This release wires automatic posting into the existing module save lifecycle for Invoices, Bills, CustomerPayments and VendorPayments.

## Transaction workflow
1. Create the invoice/bill header in Draft.
2. Add its detail rows in InvoiceLines/BillLines.
3. Set the document status to Posted (or another non-Draft, non-Cancelled status) and save.
4. The module hook posts inventory movements and balanced journal entries, then saves all changed sheets in one GitHub commit.
5. Customer/vendor payments update the linked document balance when an InvoiceID/BillID is supplied.

## Chart of Accounts prerequisites
The engine selects accounts from ChartOfAccounts by AccountType and account-name hints:
- Asset: cash/bank, accounts receivable, inventory/stock
- Revenue: sales/service/revenue
- Expense: cost of goods sold/COGS; purchase/supplies/expense (for non-inventory bill lines)
- Liability: accounts payable/payable; Output Tax/VAT Payable if sales tax is entered; Input Tax/VAT Receivable if purchase tax is entered

Account mapping is intentionally validated; missing required accounts stop posting instead of silently creating incorrect journals.

## Important limitations
- Existing invoice/bill forms are still header forms; line rows must exist in InvoiceLines/BillLines before posting.
- Payment forms must include InvoiceID/BillID to update a specific document balance. Existing form schemas may need additional linked-document selectors to make that convenient.
- The current payment schemas may not expose InvoiceID/BillID. The engine supports them when supplied, but does not guess which document to settle.
- Average cost is calculated from inventory transaction history. FIFO/LIFO layer valuation, returns, credit notes, tax-code automation, partial invoice editing/reversal, period close, bank reconciliation matching and multi-user locking are not implemented by this first integration pass.
- Posted transactions are idempotent by source type and source ID. Do not edit/delete posted financial transactions; reverse them with a correcting transaction instead.
- Test with a copy of the private database before production. This has not been tested against the user's live GitHub repository.

## Additional accounting handlers in this build
- Cost issue resolver supports FIFO, LIFO and weighted-average method labels from CompanySetup.
- SalesReturns and PurchaseReturns handlers are available when the corresponding return records provide ReturnID, ReturnNo, ReturnDate, ItemID, Quantity and Amount/TotalAmount.
- InventoryTransactions with TransactionType containing "adjustment" can post stock adjustments when the Chart of Accounts includes a suitable inventory gain/loss offset account.

These handlers are engine capabilities; existing return/adjustment UI forms and live transaction workflows still require validation against the live workbook data and field names.
