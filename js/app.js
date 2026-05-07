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

// INICIAR
mostrarProductos();

// CALCULAR TOTAL
function calcularTotal(p) {
    return p.stockInicial + p.entrada - p.salida;
}

// GUARDAR PRODUCTO
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

// MOSTRAR PRODUCTOS
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

// GUARDAR EN LOCALSTORAGE
function guardar() {
    localStorage.setItem("productos", JSON.stringify(productos));
    mostrarProductos();
}

// ELIMINAR
function eliminar(index) {
    productos.splice(index, 1);
    guardar();
}

// EDITAR
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

// BUSCADOR
buscar.addEventListener("keyup", function() {
    let texto = buscar.value.toLowerCase();

    let filtrados = productos.filter(p =>
        p.nombre.toLowerCase().includes(texto)
    );

    mostrarProductos(filtrados);
});

// CARGAR SELECT
function cargarSelect() {
    selectProducto.innerHTML = "";

    productos.forEach((p, index) => {
        selectProducto.innerHTML += `<option value="${index}">${p.nombre}</option>`;
    });
}

// AGREGAR A NOTA (SIN USAR INDEX ❗)
function agregarNota() {
    let index = selectProducto.value;
    let cantidad = parseInt(document.getElementById("cantidadNota").value);

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
        cantidad: cantidad
    });

    mostrarNota();
}

// MOSTRAR NOTA
function mostrarNota() {
    tablaNota.innerHTML = "";

    nota.forEach(n => {
        tablaNota.innerHTML += `
        <tr>
            <td>${n.nombre}</td>
            <td>${n.cantidad}</td>
        </tr>`;
    });
}

// GENERAR EXCEL + DESCONTAR INVENTARIO ✅
function generarExcelNota() {

    if (nota.length === 0) {
        alert("No hay productos en la nota");
        return;
    }

    let persona = document.getElementById("persona").value;

    if (!persona) {
        alert("Escribe el nombre");
        return;
    }

    let datos = [];

    datos.push(["LISTA DE MATERIAL DE OFICINA"]);
    datos.push(["MUNICIPIO DE CHAPULHUACÁN"]);
    datos.push([]);
    datos.push(["Recibe:", persona]);
    datos.push(["Fecha:", new Date().toLocaleDateString()]);
    datos.push([]);

    datos.push(["No.", "Producto", "Descripción", "Unidad", "Cantidad"]);

    nota.forEach((n, i) => {
        let producto = productos.find(p => p.nombre === n.nombre);

        if (producto) {
            datos.push([
                i + 1,
                producto.nombre,
                producto.descripcion || "",
                producto.unidad,
                n.cantidad
            ]);

            // 🔻 DESCONTAR BIEN
            producto.salida += n.cantidad;
        }
    });

    guardar();

    let hoja = XLSX.utils.aoa_to_sheet(datos);

    hoja['!cols'] = [
        { wch: 5 },
        { wch: 25 },
        { wch: 35 },
        { wch: 10 },
        { wch: 10 }
    ];

    let libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Nota");

    XLSX.writeFile(libro, "Nota_Material.xlsx");

    nota = [];
    mostrarNota();
}

// EXPORTAR INVENTARIO
function exportarExcel() {
    let datos = productos.map((p, i) => ({
        No: i + 1,
        Producto: p.nombre,
        Total: calcularTotal(p)
    }));

    let hoja = XLSX.utils.json_to_sheet(datos);
    let libro = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(libro, hoja, "Inventario");
    XLSX.writeFile(libro, "Inventario.xlsx");
}