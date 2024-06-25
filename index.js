import xlsxPopulate from 'xlsx-populate'


// // Forma promesa
// xlsxPopulate.fromBlankAsync()
//     .then(workbook => {
//         workbook.sheet(0).cell('A1').value('Hello world');
//         return workbook.toFileAsync("./salida.xlsx");
//     })



//Forma async await
async function main () {

    // Escribimos valores

    const workbook =  await xlsxPopulate.fromBlankAsync();

    workbook.sheet(0).cell('A1').value("Nombre")
    workbook.sheet(0).cell('B1').value("Apellido")
    workbook.sheet(0).cell('C1').value("Edad")

    workbook.sheet(0).cell('A2').value("Salu")
    workbook.sheet(0).cell('B2').value("Robles Teran")
    workbook.sheet(0).cell('C2').value("22")

    workbook.toFileAsync("./salida.xlsx");

    // Leemos valores

    const readbook = await xlsxPopulate.fromFileAsync('./salida.xlsx');
    const value = readbook.sheet('Sheet1').cell('A1').value();
    console.log(value);

    // Para leer todo

    const allValue = readbook.sheet('Sheet1').usedRange().value();

    console.log(allValue);

    // Para leer un rango
    const rangeValue = readbook.sheet('Sheet1').range('A1:B2').value();
    console.log(rangeValue);

    // Creamos varios datos

    const multiValue =  await xlsxPopulate.fromBlankAsync();
    multiValue.sheet(0).cell('A1').value([
        [new Date().getDate(), 1, 2],
        ["Nombre", "Apellido", "Edad"],
        ["Juan", "Perez", 18],
        ["Salu", "Robles", "Teran", 22],
        ["Vale", "Burgos", 24]
    ])
    multiValue.toFileAsync("./multisalida.xlsx");

    // Editar excel
    const editValue = await xlsxPopulate.fromFileAsync('./salida.xlsx');

    // ver nombre de la hoja
    const sheet = editValue.sheet(0);
    console.log(sheet.name());

    // Crear una nueva hoja
    editValue.addSheet("Hoja 2");
    editValue.toFileAsync('./salida2.xlsx')

    // Listar todas las hojas
    console.log(editValue.sheets().map(sheet => sheet.name()));

    //Renombramos una hoja
    editValue.sheet('Hoja 2').name('Hoja de prueba');
    editValue.toFileAsync('./salida3.xlsx');

    // Eliminamos una hoja
    editValue.deleteSheet('Hoja de prueba');
    editValue.toFileAsync('./salida4.xlsx');

    //ponemos contraseña
    editValue.toFileAsync('./seguro.xlsx', {
        Password: "1234"
    });
}



main();