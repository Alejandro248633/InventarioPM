let lista = document.getElementById("lista");
let formulario = document.getElementById("formulario");
let buscar = document.getElementById("buscar");
let total = document.getElementById("total");

let tablaNota = document.getElementById("tablaNota");
let selectProducto = document.getElementById("productoNota");

let productos = JSON.parse(localStorage.getItem("productos")) || [];
let nota = [];

let editando = false;
let indexEditar = null;

mostrarProductos();

function calcularTotal(p) {
    return p.stockInicial + p.entrada - p.salida;
}

formulario.addEventListener("submit", function(e) {

    e.preventDefault();

    let producto = {
        nombre: document.getElementById("nombre").value,
        descripcion: document.getElementById("descripcion").value,
        proveedor: document.getElementById("proveedor").value,
        unidad: document.getElementById("unidad").value,
        stockInicial: parseInt(document.getElementById("stockInicial").value) || 0,
        entrada: parseInt(document.getElementById("entrada").value) || 0,
        salida: parseInt(document.getElementById("salida").value) || 0
    };

    if (editando) {

        productos[indexEditar] = producto;
        editando = false;

    } else {

        productos.push(producto);
    }

    guardar();
    formulario.reset();
});

function mostrarProductos(listaProductos = productos) {

    lista.innerHTML = "";

    listaProductos.forEach((producto, index) => {

        let totalCalc = calcularTotal(producto);

        lista.innerHTML += `
        <tr>
            <td>${index + 1}</td>
            <td>${producto.nombre}</td>
            <td>${producto.descripcion || ""}</td>
            <td>${producto.proveedor || ""}</td>
            <td>${producto.unidad}</td>
            <td>${producto.stockInicial}</td>
            <td>${producto.entrada}</td>
            <td>${producto.salida}</td>
            <td>${totalCalc}</td>
            <td>
                <button class="editar" onclick="editar(${index})">Editar</button>
                <button class="eliminar" onclick="eliminar(${index})">Eliminar</button>
            </td>
        </tr>`;
    });

    total.textContent = productos.length;
    cargarSelect();
}

function guardar() {

    localStorage.setItem("productos", JSON.stringify(productos));
    mostrarProductos();
}

function eliminar(index) {

    productos.splice(index, 1);
    guardar();
}

function editar(index) {

    let p = productos[index];

    document.getElementById("nombre").value = p.nombre;
    document.getElementById("descripcion").value = p.descripcion;
    document.getElementById("proveedor").value = p.proveedor;
    document.getElementById("unidad").value = p.unidad;
    document.getElementById("stockInicial").value = p.stockInicial;
    document.getElementById("entrada").value = p.entrada;
    document.getElementById("salida").value = p.salida;

    editando = true;
    indexEditar = index;
}

buscar.addEventListener("keyup", function() {

    let texto = buscar.value.toLowerCase();

    let filtrados = productos.filter(p =>
        p.nombre.toLowerCase().includes(texto)
    );

    mostrarProductos(filtrados);
});

function cargarSelect() {

    selectProducto.innerHTML = "";

    productos.forEach((p, index) => {

        selectProducto.innerHTML += `
        <option value="${index}">
            ${p.nombre}
        </option>`;
    });
}

function agregarNota() {

    let index = selectProducto.value;
    let cantidad = parseInt(document.getElementById("cantidadNota").value);
    let codigo = document.getElementById("codigoProducto").value;
    let precioUnitario = parseFloat(document.getElementById("precioUnitario").value) || 0;

    if (!cantidad || cantidad <= 0) {

        alert("Cantidad inválida");
        return;
    }

    let producto = productos[index];

    if (calcularTotal(producto) < cantidad) {

        alert("No hay suficiente stock");
        return;
    }

    nota.push({
        nombre: producto.nombre,
        codigo: codigo,
        cantidad: cantidad,
        precioUnitario: precioUnitario,
        importe: cantidad * precioUnitario
    });

    document.getElementById("codigoProducto").value = "";
    document.getElementById("precioUnitario").value = "";
    document.getElementById("cantidadNota").value = "";

    mostrarNota();
}

function mostrarNota() {

    tablaNota.innerHTML = "";

    nota.forEach(n => {

        tablaNota.innerHTML += `
        <tr>
            <td>${n.codigo}</td>
            <td>${n.nombre}</td>
            <td>${n.cantidad}</td>
            <td>$${n.precioUnitario.toFixed(2)}</td>
            <td>$${n.importe.toFixed(2)}</td>
        </tr>`;
    });
}

/* CARGAR IMÁGENES */
function cargarImagen(url) {

    return new Promise((resolve, reject) => {

        const xhr = new XMLHttpRequest();

        xhr.open("GET", url, true);

        xhr.responseType = "blob";

        xhr.onload = function () {

            if (this.status === 200) {

                const reader = new FileReader();

                reader.onload = function () {

                    resolve(reader.result);
                };

                reader.readAsDataURL(this.response);

            } else {

                reject("Error cargando imagen");
            }
        };

        xhr.onerror = function () {

            reject("No se pudo cargar la imagen");
        };

        xhr.send();
    });
}

/* GENERAR NOTA EXCEL */
async function generarExcelNota() {

    if (nota.length === 0) {

        alert("No hay productos en la nota");
        return;
    }

    let persona = document.getElementById("persona").value;

    if (!persona) {

        alert("Escribe el nombre");
        return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("NOTA");

    worksheet.pageSetup.orientation = 'landscape';
    worksheet.columns = [
        { width: 12 },
        { width: 20 },
        { width: 25 },
        { width: 15 },
        { width: 12 },
        { width: 15 }
    ];

    /* LOGOS */

    try {
        const logo1 = await cargarImagen("../img/logo1.png");
        const img1 = workbook.addImage({ base64: logo1, extension: 'png' });
        worksheet.addImage(img1, { tl: { col: 0.2, row: 0.2 }, ext: { width: 80, height: 60 } });
    } catch (e) { console.log("Logo 1 no disponible"); }

    try {
        const logo2 = await cargarImagen("../img/logo2.png");
        const img2 = workbook.addImage({ base64: logo2, extension: 'png' });
        worksheet.addImage(img2, { tl: { col: 1.5, row: 0.1 }, ext: { width: 300, height: 90 } });
    } catch (e) { console.log("Logo 2 no disponible"); }

    try {
        const logo3 = await cargarImagen("../img/logo3.png");
        const img3 = workbook.addImage({ base64: logo3, extension: 'png' });
        worksheet.addImage(img3, { tl: { col: 5.2, row: 0.2 }, ext: { width: 80, height: 60 } });
    } catch (e) { console.log("Logo 3 no disponible"); }

    /* TÍTULO */
    worksheet.mergeCells('A6:F6');
    worksheet.getCell('A6').value = "NOTA DE ENTREGA";
    worksheet.getCell('A6').font = { size: 16, bold: true };
    worksheet.getCell('A6').alignment = { horizontal: 'center' };
    worksheet.getCell('A6').fill = { type: "pattern", pattern: "solid", fgColor: { argb: "C00000" } };
    worksheet.getCell('A6').font.color = { argb: "FFFFFFFF" };

    /* SUBTÍTULO */
    worksheet.mergeCells('A7:F7');
    worksheet.getCell('A7').value = "Comprobante de Entrega de Material";
    worksheet.getCell('A7').alignment = { horizontal: 'center' };
    worksheet.getCell('A7').font = { italic: true };

    /* INFORMACIÓN DEL RESPONSABLE */
    worksheet.mergeCells('A9:B9');
    worksheet.getCell('A9').value = "INFORMACIÓN DEL RESPONSABLE";
    worksheet.getCell('A9').font = { bold: true };
    worksheet.getCell('A9').fill = { type: "pattern", pattern: "solid", fgColor: { argb: "C00000" } };
    worksheet.getCell('A9').font.color = { argb: "FFFFFFFF" };

    worksheet.mergeCells('D9:F9');
    worksheet.getCell('D9').value = "INFORMACIÓN DEL QUE RECIBE";
    worksheet.getCell('D9').font = { bold: true };
    worksheet.getCell('D9').fill = { type: "pattern", pattern: "solid", fgColor: { argb: "C00000" } };
    worksheet.getCell('D9').font.color = { argb: "FFFFFFFF" };

    let responsableTitular = document.getElementById("responsableTitular").value || "";
    let responsableArea = document.getElementById("responsableArea").value || "";
    let responsableTelefono = document.getElementById("responsableTelefono").value || "";
    let responsableEmail = document.getElementById("responsableEmail").value || "";
    let areaRecibe = document.getElementById("areaRecibe").value || "";
    let partida = document.getElementById("partida").value || "";
    let solicitud = document.getElementById("solicitud").value || "";

    worksheet.getCell('A10').value = "Titular:";
    worksheet.getCell('A10').font = { bold: true };
    worksheet.mergeCells('B10:C10');
    worksheet.getCell('B10').value = responsableTitular;

    worksheet.getCell('D10').value = "Titular:";
    worksheet.getCell('D10').font = { bold: true };
    worksheet.mergeCells('E10:F10');
    worksheet.getCell('E10').value = persona;

    worksheet.getCell('A11').value = "Área:";
    worksheet.getCell('A11').font = { bold: true };
    worksheet.mergeCells('B11:C11');
    worksheet.getCell('B11').value = responsableArea;

    worksheet.getCell('D11').value = "Área:";
    worksheet.getCell('D11').font = { bold: true };
    worksheet.mergeCells('E11:F11');
    worksheet.getCell('E11').value = areaRecibe;

    worksheet.getCell('A12').value = "Teléfono:";
    worksheet.getCell('A12').font = { bold: true };
    worksheet.mergeCells('B12:C12');
    worksheet.getCell('B12').value = responsableTelefono;

    worksheet.getCell('D12').value = "Partida Presupuestal:";
    worksheet.getCell('D12').font = { bold: true };
    worksheet.mergeCells('E12:F12');
    worksheet.getCell('E12').value = partida;

    worksheet.getCell('A13').value = "Email:";
    worksheet.getCell('A13').font = { bold: true };
    worksheet.mergeCells('B13:C13');
    worksheet.getCell('B13').value = responsableEmail;

    worksheet.getCell('D13').value = "Solicitud:";
    worksheet.getCell('D13').font = { bold: true };
    worksheet.mergeCells('E13:F13');
    worksheet.getCell('E13').value = solicitud;

    /* DATOS DE ENTREGA */
    worksheet.mergeCells('A15:F15');
    worksheet.getCell('A15').value = "DATOS DE ENTREGA";
    worksheet.getCell('A15').font = { bold: true };
    worksheet.getCell('A15').fill = { type: "pattern", pattern: "solid", fgColor: { argb: "C00000" } };
    worksheet.getCell('A15').font.color = { argb: "FFFFFFFF" };

    let numeroNota = document.getElementById("numeroNota").value || "";
    let fechaEmision = document.getElementById("fechaEmision").value || "";
    let numeroPedido = document.getElementById("numeroPedido").value || "";
    let fechaEntrega = document.getElementById("fechaEntrega").value || "";
    let numeroFactura = document.getElementById("numeroFactura").value || "";
    let horarioEntrega = document.getElementById("horarioEntrega").value || "";

    worksheet.getCell('A16').value = "N° de Nota:";
    worksheet.getCell('A16').font = { bold: true };
    worksheet.mergeCells('B16:C16');
    worksheet.getCell('B16').value = numeroNota;

    worksheet.getCell('D16').value = "Fecha de Emisión:";
    worksheet.getCell('D16').font = { bold: true };
    worksheet.mergeCells('E16:F16');
    worksheet.getCell('E16').value = fechaEmision;

    worksheet.getCell('A17').value = "N° de Pedido:";
    worksheet.getCell('A17').font = { bold: true };
    worksheet.mergeCells('B17:C17');
    worksheet.getCell('B17').value = numeroPedido;

    worksheet.getCell('D17').value = "Fecha de Entrega:";
    worksheet.getCell('D17').font = { bold: true };
    worksheet.mergeCells('E17:F17');
    worksheet.getCell('E17').value = fechaEntrega;

    worksheet.getCell('A18').value = "N° de Factura:";
    worksheet.getCell('A18').font = { bold: true };
    worksheet.mergeCells('B18:C18');
    worksheet.getCell('B18').value = numeroFactura;

    worksheet.getCell('D18').value = "Horario Entrega:";
    worksheet.getCell('D18').font = { bold: true };
    worksheet.mergeCells('E18:F18');
    worksheet.getCell('E18').value = horarioEntrega;

    /* TABLA DE PRODUCTOS */
    const encabezados = ["CÓDIGO", "PRODUCTO", "DESCRIPCIÓN", "UNIDAD", "CANTIDAD", "PRECIO UNITARIO", "IMPORTE"];

    let filaInicio = 20;
    const headerRow = worksheet.getRow(filaInicio);
    
    encabezados.forEach((titulo, index) => {
        const cell = headerRow.getCell(index + 1);
        cell.value = titulo;
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1F4E78" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    });

    let totalImporte = 0;

    /* PRODUCTOS */
    nota.forEach((n, i) => {
        let producto = productos.find(p => p.nombre === n.nombre);
        let fila = worksheet.getRow(filaInicio + i + 1);

        fila.getCell(1).value = n.codigo;
        fila.getCell(2).value = n.nombre;
        fila.getCell(3).value = producto.descripcion || "";
        fila.getCell(4).value = producto.unidad;
        fila.getCell(5).value = n.cantidad;
        fila.getCell(6).value = n.precioUnitario;
        fila.getCell(7).value = n.importe;

        totalImporte += n.importe;

        fila.eachCell(cell => {
            cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
            cell.alignment = { horizontal: "center", vertical: "middle" };
        });

        /* DESCONTAR INVENTARIO */
        producto.salida += n.cantidad;
    });

    guardar();

    /* TOTAL IMPORTE */
    let filaTotal = filaInicio + nota.length + 1;
    worksheet.mergeCells(`A${filaTotal}:F${filaTotal}`);
    worksheet.getCell(`A${filaTotal}`).value = `TOTAL IMPORTE: $${totalImporte.toFixed(2)}`;
    worksheet.getCell(`A${filaTotal}`).font = { bold: true };
    worksheet.getCell(`A${filaTotal}`).alignment = { horizontal: "right" };
    worksheet.getCell(`A${filaTotal}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9D9D9" } };

    /* FIRMAS */
    let filaFirmas = filaTotal + 3;

    worksheet.mergeCells(`A${filaFirmas}:B${filaFirmas}`);
    worksheet.mergeCells(`D${filaFirmas}:E${filaFirmas}`);

    worksheet.getCell(`A${filaFirmas}`).value = "ENTREGA";
    worksheet.getCell(`D${filaFirmas}`).value = "RECIBE";

    worksheet.getCell(`A${filaFirmas}`).font = { bold: true };
    worksheet.getCell(`D${filaFirmas}`).font = { bold: true };
    worksheet.getCell(`A${filaFirmas}`).alignment = { horizontal: "center" };
    worksheet.getCell(`D${filaFirmas}`).alignment = { horizontal: "center" };

    /* DESCARGAR */
    const buffer = await workbook.xlsx.writeBuffer();

    saveAs(
        new Blob([buffer], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }),
        "NOTA_DE_ENTREGA.xlsx"
    );

    nota = [];
    mostrarNota();
}

/* EXPORTAR INVENTARIO */

function exportarExcel() {

    let datos = productos.map((p, i) => ({
        No: i + 1,
        Producto: p.nombre,
        Descripcion: p.descripcion,
        Proveedor: p.proveedor,
        Unidad: p.unidad,
        Stock: calcularTotal(p)
    }));

    let hoja = XLSX.utils.json_to_sheet(datos);

    let libro = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(libro, hoja, "Inventario");

    XLSX.writeFile(libro, "Inventario.xlsx");
}
