const excel = require('exceljs');
const path = require('path');
let ruta = path.join(__dirname.split('excel_controller')[0] + "Informes")
console.log(ruta)
// EXCEL/src/ [0]
// excel [1]

const meses = [
    '',
    'Enero',
    'Febrero',
    'Marzo',
    'Abril',
    'Mayo',
    'Junio',
    'Julio',
    'Agosto',
    'Septiembre',
    'Octubre',
    'Noviembre',
    'Diciembre',
];

const fecha = new Date();

function dia() {
    return fecha.getDate();
}

function mes() {
    return meses[fecha.getUTCMonth() + 1];
}

function year() {
    return fecha.getUtcFullYear();
}

function crear(datos, nombre_espacio_trabajo = "Reporte") {

    let libro_de_trabajo = new excel.Workbook();
    let estilos_de_trabajo = libro_de_trabajo.addWorksheet(nombre_espacio_trabajo);

    // Titulo
    // mergeCells combina una celda hasta otra celda
    // Example: mergeCells("A1", "C1")
    estilos_de_trabajo.mergeCells('B1', 'E4');
    /**
     * getCell("B1") nos permite agregarle tanto texto como otro valor con la propiedad value
     * tambien podemos utilizar la propiedad de style para poder darle unos estilos a la celda
     */
    estilos_de_trabajo.getCell("B1").value = "Titulo";
    estilos_de_trabajo.getCell("B1").style = {
        alignment: {
            horizontal: 'center',
            vertical: 'middle'
        },
        font: {
            bold: true
        }
    };

    // Fecha de la descarga
    estilosDeTrabajo.mergeCells('B6', 'E6');
    estilosDeTrabajo.getCell('B6').value = `${dia()} de ${mes()} del ${year()}`;
    estilosDeTrabajo.getCell('B6').style = {
        alignment: {
            horizontal: 'center',
            vertical: 'middle',
        },
        font: {
            bold: true,
        },
    };

    // Inicio del esquema de datos para rellenar
    // Utilizamos el getRow para empezar la maqueta donde pondremos nuestra lista de datos
    // Cabe destacar ver que no entre en conflicto con el lo hecho anteriormente con el titulo
    // Por que sino se tendri que cambiar la pospicion del get Row para empezar a meter toda la data en la maqueta
    estilosDeTrabajo.getRow(9).values = ['Nombre', 'Descripción', 'Vigente'];
    // Las columns nos permiten estructurar que datos entraran en las celdas en base a su key 
    // podremos agregarle funcionalidades y estilos entre otros
    estilosDeTrabajo.columns = [
        { key: 'Nombre', width: 45 },
        { key: 'Descripcion', width: 45 },
        { key: 'Vigencia', width: 35 },
    ];
    // addRows nos permite ingresar toda la lista de datos
    estilosDeTrabajo.addRows(datos);
    // Este es un filtrador
    estilosDeTrabajo.autoFilter = {
        from: 'A9',
        to: 'C9',
    };

    return libro_de_trabajo.xlsx.writeFile(`${ruta}/test.xlsx`)
}