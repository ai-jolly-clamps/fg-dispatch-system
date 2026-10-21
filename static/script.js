// ======================================
// FG Dispatch System
// script.js
// ======================================

console.log("FG Dispatch System Loaded");

// ======================================
// SALES PAGE
// ADD ROW
// ======================================

function addRow() {

    let tableBody =
        document.getElementById("tableBody");

    if (!tableBody) return;

    let row = document.createElement("tr");

    row.innerHTML = `

        <td>
            <input type="text"
                   class="form-control m_code">
        </td>

        <td>
            <input type="text"
                   class="form-control customer">
        </td>

        <td>
            <input type="text"
                   class="form-control part_number">
        </td>

        <td>
            <input type="text"
                   class="form-control size">
        </td>

        <td>
            <input type="number"
                   class="form-control dispatch_qty">
        </td>

        <td>
            <input type="text"
                   class="form-control invoice_no">
        </td>

        <td>
            <input type="text"
                   class="form-control remarks">
        </td>

        <td>
            <button type="button"
                    class="btn btn-danger"
                    onclick="deleteRow(this)">
                Delete
            </button>
        </td>

    `;

    tableBody.appendChild(row);
}

// ======================================
// DELETE ROW
// ======================================

function deleteRow(button) {

    let row = button.closest("tr");

    if (row) {
        row.remove();
    }
}

// ======================================
// VALIDATE SALES DATA
// ======================================

function validateSalesRows() {

    let rows =
        document.querySelectorAll("#tableBody tr");

    if (rows.length === 0) {

        alert("Please add at least one row.");
        return false;
    }

    for (let row of rows) {

        let m_code =
            row.querySelector(".m_code")?.value.trim();

        let customer =
            row.querySelector(".customer")?.value.trim();

        let qty =
            row.querySelector(".dispatch_qty")?.value.trim();

        if (
            m_code === "" ||
            customer === "" ||
            qty === ""
        ) {

            alert(
                "M Code, Customer and Dispatch Qty are mandatory."
            );

            return false;
        }
    }

    return true;
}

// ======================================
// SALES SAVE
// ======================================

function saveAllRows()
{
    let dispatchDate =
        document.getElementById("dispatch_date").value;

    if(dispatchDate === "")
    {
        alert("Please select Dispatch Date");
        return;
    }

    let rows = [];

    document.querySelectorAll("#tableBody tr")
    .forEach(function(row){

        let mcode =
        row.querySelector(".m_code").value;

        let customer =
        row.querySelector(".customer").value;

        let part =
        row.querySelector(".part_number").value;

        let size =
        row.querySelector(".size").value;

        let qty =
        row.querySelector(".dispatch_qty").value;

        let invoice =
        row.querySelector(".invoice_no").value;

        let remarks =
        row.querySelector(".remarks").value;

        if(mcode !== "")
        {
            rows.push({

                m_code: mcode,
                customer: customer,
                part_number: part,
                size: size,
                dispatch_qty: qty,
                invoice_no: invoice,
                remarks: remarks

            });
        }

    });

    if(rows.length === 0)
    {
        alert("No Data Found");
        return;
    }

    fetch("/save_sales_bulk", {

        method: "POST",

        headers: {
            "Content-Type": "application/json"
        },

        body: JSON.stringify({

            dispatch_date: dispatchDate,
            rows: rows

        })

    })

    .then(response => response.json())

    .then(data => {

        if(data.success)
        {
            alert("Data Saved Successfully");

            document.getElementById("tableBody").innerHTML = "";

            addRow();
        }
        else
        {
            alert(data.error);
        }

    })

    .catch(error => {

        console.log(error);

        alert("Save Failed");

    });
}




// ======================================
// AUTO REFRESH FG TABLE
// ======================================

function refreshFGTable() {

    let fgTable =
        document.getElementById("fgTable");

    if (fgTable) {

        console.log(
            "Refreshing FG Pending Records..."
        );

        loadData();
    }
}

// Every 60 seconds

setInterval(function () {

    let fgTable =
        document.getElementById("fgTable");

    if (fgTable) {

        refreshFGTable();

    }

}, 60000);

// ======================================
// PAGE READY
// ======================================

document.addEventListener(
    "DOMContentLoaded",
    function () {

        console.log(
            "Application Ready"
        );

    }
);