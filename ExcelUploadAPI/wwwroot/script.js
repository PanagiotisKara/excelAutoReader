        const fileInput = document.getElementById("fileInput");
        const fileNameDisplay = document.getElementById("fileName");
        const uploadBtn = document.getElementById("uploadBtn");

        fileInput.addEventListener("change", function () {
            if (fileInput.files.length > 0) {
                const fileName = fileInput.files[0].name;

                if (!fileName.endsWith(".xlsx") && !fileName.endsWith(".xls")) {
                    fileNameDisplay.innerHTML = "❌ Invalid file type. Please select an Excel file (.xlsx or .xls)";
                    uploadBtn.style.display = "none";
                    return;
                }

                fileNameDisplay.innerHTML = `📂 Selected: ${fileName}`;
                uploadBtn.style.display = "inline-block";
            }
        });

        function uploadExcel() {
    let file = document.getElementById("fileInput").files[0];
    if (!file) {
        alert("Please select a file!");
        return;
    }

    let formData = new FormData();
    formData.append("file", file);

    fetch("http://localhost:5143/api/excel/upload", {
        method: "POST",
        body: formData
    })
    .then(response => response.json())
    .then(responseData => {
        console.log("API Response:", responseData);

        let tableContainer = document.getElementById("dataTableContainer");
        let tableBody = document.querySelector("#dataTable tbody");
        tableBody.innerHTML = ""; 

        if (!responseData || !responseData.data || responseData.data.length === 0) {
            tableContainer.style.display = "none"; 
            tableBody.innerHTML = "<tr><td colspan='4'>No data found in the file.</td></tr>";
            return;
        }

        tableContainer.style.display = "block";

        responseData.data.forEach((item, index) => {
            console.log("Processing item:", item);
            console.log("Object keys:", Object.keys(item));

            let row = `<tr>
                <td>${index + 1}</td>
                <td>${item.crmbatrid || "N/A"}</td>
                <td>${item.preInvoicedocNo || "N/A"}</td>
                <td>${item.baTR_Amount || "N/A"}</td>  
                <td>${item.notes || "N/A"}</td>       
            </tr>`;
            tableBody.innerHTML += row;
        });
    })
    .catch(error => {
        console.error("Fetch error:", error);
        alert("Failed to upload the file. Make sure the server is running.");
    });
}