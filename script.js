let hot; // Instancia de Handsontable
let fileName = "archivo_editado.xlsx";

document.getElementById('file-input').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    fileName = file.name;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Tomamos la primera hoja
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convertir a formato JSON (matriz de matrices)
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        renderTable(jsonData);
        document.getElementById('save-btn').style.display = 'inline-block';
    };
    reader.readAsArrayBuffer(file);
});

function renderTable(data) {
    const container = document.getElementById('excel-table');
    
    // Si ya existe una tabla, la destruimos para cargar la nueva
    if (hot) hot.destroy();

    hot = new Handsontable(container, {
        data: data,
        rowHeaders: true,
        colHeaders: true,
        contextMenu: true,
        licenseKey: 'non-commercial-and-evaluation', // Para uso educativo/personal
        width: '100%',
        height: 'auto',
        stretchH: 'all'
    });
}

document.getElementById('save-btn').addEventListener('click', function() {
    if (!hot) return;

    // Obtener los datos actuales de la tabla
    const editedData = hot.getData();

    // Crear un nuevo libro de trabajo y hoja
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(editedData);

    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Hoja1");

    // Generar el archivo y disparar la descarga
    XLSX.writeFile(newWorkbook, fileName);
});
