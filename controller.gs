//GETTING API SECRET KEY FROM ANOTHER SCRIPT
const API_KEY = getAPIKey();

const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getActiveSheet();
const lastRowSortMap = 'V411';
let queryParam = 'xxx';

// custom menu
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    //api gathering data menu
    ui.createMenu('🎬 Ver')
        .addItem('🙈 Ocultar Info', 'hideInfo')
        .addItem('👀 Mostrar Info', 'showInfo')
        .addToUi();
    ui.createMenu('🎬 Datos')
        .addItem('♾️ Actualizar TODO', 'updateData')
        .addItem('1️⃣ Actualizar Fila Seleccionada', 'updateSelected')
        .addToUi();
    //sorting menu -> TO CHANGE: SORT BY SELECTED COLUMN MENU ASC DESC
    ui.createMenu('🎬 Ordenar')
        .addSubMenu(ui.createMenu('Nuestra Nota Media')
            .addItem('🏆 Mejores', 'sortByAverageDesc')
            .addItem('🤮 Peores', 'sortByAverageAsc'))
        .addSubMenu(ui.createMenu('Mico')
            .addItem('🏆 Mejores', 'sortByMicoDesc')
            .addItem('🤮 Peores', 'sortByMicoAsc'))
        .addSubMenu(ui.createMenu('Gimi')
            .addItem('🏆 Mejores', 'sortByGimiDesc')
            .addItem('🤮 Peores', 'sortByGimiAsc'))
        .addSubMenu(ui.createMenu('Masip')
            .addItem('🏆 Mejores', 'sortByMasipDesc')
            .addItem('🤮 Peores', 'sortByMasipAsc'))
        .addSubMenu(ui.createMenu('Cauda')
            .addItem('🏆 Mejores', 'sortByCaudaDesc')
            .addItem('🤮 Peores', 'sortByCaudaAsc'))
        .addSubMenu(ui.createMenu('IMDB')
            .addItem('🏆 Mejores', 'sortByIMDBDesc')
            .addItem('🤮 Peores', 'sortByIMDBAsc'))
        .addToUi();

}

function hideInfo() {
    sheet.hideColumns(2, 7);
}

function showInfo() {
    sheet.showColumns(2, 7);
}
//////  SORTING METHODS ///////////
//Average
function sortByAverageDesc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    //range.sort({column: 19, ascending: false});
    range.sort([{
        column: 19,
        ascending: false
    }, {
        column: 8,
        ascending: false
    }]);
    toastMessageTitle('¡Hoja ordenada! Índice: Nota media. Orden: Descendente.', '¡Éxito!');
}

function sortByAverageAsc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    range.sort([{
        column: 19,
        ascending: true
    }, {
        column: 8,
        ascending: true
    }]);
    toastMessageTitle('¡Hoja ordenada! Índice: Nota media. Orden: Ascendente.', '¡Éxito!');
}
//Marcos
function sortByMicoDesc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    range.sort({
        column: 9,
        ascending: false
    });
    toastMessageTitle('¡Hoja ordenada! Índice: MICO. Orden: Descendente.', '¡Éxito!');
}

function sortByMicoAsc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    range.sort({
        column: 9,
        ascending: true
    });
    toastMessageTitle('¡Hoja ordenada! Índice: MICO. Orden: Ascendente.', '¡Éxito!');
}
//Gimi
function sortByGimiDesc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    range.sort({
        column: 10,
        ascending: false
    });
    toastMessageTitle('¡Hoja ordenada! Índice: GIMI. Orden: Descendente.', '¡Éxito!');
}

function sortByGimiAsc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    range.sort({
        column: 10,
        ascending: true
    });
    toastMessageTitle('¡Hoja ordenada! Índice: GIMI. Orden: Ascendente.', '¡Éxito!');
}
//Masip
function sortByMasipDesc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    range.sort({
        column: 11,
        ascending: false
    });
    toastMessageTitle('¡Hoja ordenada! Índice: MASIP. Orden: Descendente.', '¡Éxito!');
}

function sortByMasipAsc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    range.sort({
        column: 11,
        ascending: true
    });
    toastMessageTitle('¡Hoja ordenada! Índice: MASIP. Orden: Ascendente.', '¡Éxito!');
}
//Cauda
function sortByCaudaDesc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    range.sort({
        column: 18,
        ascending: false
    });
    toastMessageTitle('¡Hoja ordenada! Índice: CAUDA. Orden: Descendente.', '¡Éxito!');
}

function sortByCaudaAsc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    range.sort({
        column: 18,
        ascending: true
    });
    toastMessageTitle('¡Hoja ordenada! Índice: CAUDA. Orden: Ascendente.', '¡Éxito!');
}
//IMDB
function sortByIMDBDesc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    range.sort({
        column: 8,
        ascending: false
    });
    toastMessageTitle('¡Hoja ordenada! Índice: IMDB. Orden: Descendente.', '¡Éxito!');
}

function sortByIMDBAsc() {
    var range = sheet.getRange("A3:" + lastRowSortMap);
    range.sort({
        column: 8,
        ascending: true
    });
    toastMessageTitle('¡Hoja ordenada! Índice: IMDB. Orden: Ascendente.', '¡Éxito!');
}

function addNewMovie() {
    var ui = SpreadsheetApp.getUi();
    var index = 3;
    var result = ui.prompt("Nombre de la película");

    if (result.getResponseText() != "") {
        sheet.insertRowBefore(index);
        toastMessageTitle('"' + result.getResponseText() + '"' + ' actualizado correctamente.', '¡Éxito!');
        sheet.getRange(index, 1).setValue(result.getResponseText());
        var searchKey = sheet.getRange(index, 1).getValue();
        if (searchKey != '' || searchKey != '-') {
            sheet.getRange(index, 1).setBackground('#CCCCCC');
            let movieInfo = isMovieOrTV(searchKey);

            //-- SETTING VALUES --//  
            if (movieInfo !== null) {
                for (let i = 0; i < movieInfo.length; i++) {
                    if (i == 1) {
                        sheet.getRange(index, i + 2).setNote(movieInfo[i]);
                    } else {
                        sheet.getRange(index, i + 2).setValue(movieInfo[i])
                    }
                }
            }
            toastMessageTitle('"' + searchKey + '"' + ' actualizado correctamente.', '¡Éxito!');
        }
    }

}

///////////////////////////////////////////////////////////////////
////////////////API DATA GATHERING/////////1.0/////////////////////
///////////////////////////////////////////////////////////////////

//////TOAST HELPER/////
function toastMessageTitle(message, title) {
    SpreadsheetApp.getActive().toast(message, title);
}

//update selected row JUST 1
function updateSelected() {
    var index = sheet.getActiveRange().getRowIndex();
    var searchKey = sheet.getRange(index, 1).getValue();
    if (searchKey != '' || searchKey != '-') {
        sheet.getRange(index, 1).setBackground('#CCCCCC');
        let movieInfo = isMovieOrTV(searchKey);

        //-- SETTING VALUES --//  
        if (movieInfo !== null) {
            for (let i = 0; i < movieInfo.length; i++) {
                if (i == 1) {
                    sheet.getRange(index, i + 2).setNote(movieInfo[i]);
                } else {
                    sheet.getRange(index, i + 2).setValue(movieInfo[i])
                }
            }
        }
        toastMessageTitle('"' + searchKey + '"' + ' actualizado correctamente.', '¡Éxito!');
    }

}
/////UPDATE SINCE SELECTED ROW
function updateData() {
    //const var
    var iterator = sheet.getActiveRange().getRowIndex();

    //main sheet iterator
    var listString;
    while (true) {
        sheet.getRange(iterator, 1).setBackground('#CCCCCC');
        sheet.getRange(iterator, 1).setBackground('green');
        var searchKey = sheet.getRange(iterator, 1).getValue();
        Utilities.sleep(10);
        searchKey = String(searchKey);
        listString = listString + searchKey + ', ';

        //while BREAK condition
        if (searchKey == '-') {
            break;
        } else if (searchKey != '') {
            let movieInfo = isMovieOrTV(searchKey);

            //-- SETTING VALUES --//
            if (movieInfo !== null) {
                for (let i = 0; i < movieInfo.length; i++) {
                    if (i == 1) {
                        sheet.getRange(iterator, i + 2).setNote(movieInfo[i]);
                    } else {
                        sheet.getRange(iterator, i + 2).setValue(movieInfo[i])
                    }
                }
            }
        } else {}
        sheet.getRange(iterator, 1).setBackground('#CCCCCC');
        iterator++;
        toastMessageTitle('"' + searchKey + '"' + ' actualizado correctamente.', '¡Éxito!');
    }

}

//   [MOVIE] - - - [SERIE]
function isMovieOrTV(searchKey) {
    let queryInfo = [];
    if (sheet.getName().includes('PELIS')) {
        queryParam = 'movie';
        queryInfo = updateMovieData(searchKey);

    } else if (sheet.getName().includes('SERIES')) {
        queryParam = 'tv';
        queryInfo = updateMovieData(searchKey);
    } else {
        toastMessageTitle('"' + searchKey + '"' + ' ERROR EN - var sheet = sheet.getName() PELIS / SERIES.', '¡ERROR!');
    }
    return queryInfo;
}

// GET MOVIE ID
function updateMovieData(searchKey) {
    var replaced = searchKey.split(' ').join('+');
    let search = 'https://api.themoviedb.org/3/search/' + queryParam + '?api_key=' + API_KEY + '&query=' + replaced + '&language=es';
    var response = UrlFetchApp.fetch(search);
    var json = response.getContentText();
    var data = JSON.parse(json);

    if (json.includes("id")) {
        return getMovieDetail(data["results"][0]["id"]);
    } else {
        let movieInfo = [];
        for (let i = 0; i < 6; i++) {
            movieInfo.push('Not found');
        }
        return movieInfo;
    }
}
//GET MOVIE DATA BY ID + APPEND CREDITS
function getMovieDetail(id) {
    var search = 'https://api.themoviedb.org/3/' + queryParam + '/' + id + '?api_key=' + API_KEY + '&append_to_response=credits' + '&language=es';
    var response = UrlFetchApp.fetch(search);
    var json = response.getContentText();
    var data = JSON.parse(json);

    var DATA_Genero = 'Not found'
    var DATA_Direccion = 'Not found';
    var DATA_Sinopsis = 'Not found';
    var DATA_Plataforma = 'Not found';
    var DATA_Duracion = 'Not found';
    var DATA_Año = 'Not found';
    var DATA_Imdb = 'Not found';

    if (Object.keys(response).length > 0) {
        //-- GENERO --//
        if (Object.keys(data["genres"]).length >= 2) {
            DATA_Genero = data["genres"][0]["name"] + ', ' + data["genres"][1]["name"];
        } else if (Object.keys(data["genres"]).length == 1) {
            DATA_Genero = data["genres"][0]["name"];
        }

        //-- SINOPSIS --//
        DATA_Sinopsis = data["overview"];
        //-- PLATAFORMA --//
        DATA_Plataforma = getPlataforma(id);
        //-- DIRECTOR --//
        if (sheet.getName().includes('PELIS')) {
            let directors = [];
            for (let i = 0; i < Object.keys(data["credits"]["crew"]).length; i++) {
                if (data["credits"]["crew"][i]["known_for_department"] == 'Directing') {
                    if ("name" in data["credits"]["crew"][i]) {
                        var director = data["credits"]["crew"][i]["name"];
                        directors.push(director);
                    }
                }
            }
            if (directors.length > 0) {
                if (directors[0] == directors[1] || directors.length == 1) {
                    DATA_Direccion = directors[0];
                } else {
                    DATA_Direccion = directors[0] + ' y ' + directors[1];
                }
            }
        } else if (sheet.getName().includes('SERIES')) {
            let duracion = data["episode_run_time"];
            let average = duracion.reduce((a, b) => a + b, 0) / duracion.length;
            let round = Math.ceil(average / 5) * 5;
            let episodios = data["number_of_episodes"];
            let temporadas = data["number_of_seasons"];
            DATA_Direccion = temporadas + ' temp  - ' + episodios + ' cap  - ' + round + ' min';

        }

        /*
            if(Object.keys(data["credits"]["crew"]).length>=2){
              if(data["credits"]["crew"][0]["name"] == data["credits"]["crew"][1]["name"]){
                DATA_Direccion = data["credits"]["crew"][0]["name"];
              }else{
                DATA_Direccion = data["credits"]["crew"][0]["name"] + 'y ' + data["credits"]["crew"][1]["name"];
              }
            }else if(Object.keys(data["credits"]["crew"]).length==1){
              DATA_Direccion = data["credits"]["crew"][0]["name"];
            }
        */

        //-- DURACION --//
        if (sheet.getName().includes('PELIS')) {
            DATA_Duracion = integerToHours(data["runtime"]);
        } else if (sheet.getName().includes('SERIES')) {
            DATA_Duracion = data["status"];
            DATA_Duracion.includes('Ended') ? DATA_Duracion = '✅ Si' : DATA_Duracion = '❌ No';
        }

        //-- AÑO    SERIES= EN EMISION --/
        if (sheet.getName().includes('PELIS')) {
            DATA_Año = data["release_date"].substring(0, 4);
        } else if (sheet.getName().includes('SERIES')) {
            DATA_Año = data["first_air_date"].substring(0, 4);
        }


        //-- IMDB --//
        DATA_Imdb = data["vote_average"];

        //-- RETURN VALUES TO MAIN ARRAY--//

        let movieInfo = [];
        movieInfo.push(DATA_Genero);
        movieInfo.push(DATA_Sinopsis);
        movieInfo.push(DATA_Direccion);
        movieInfo.push(DATA_Plataforma);
        movieInfo.push(DATA_Duracion);
        movieInfo.push(DATA_Año);
        movieInfo.push(DATA_Imdb);
        return movieInfo;
    }


}
//API CALL GET STREAM SERVICE BY MOVIE ID
function getPlataforma(id) {

    //var id = 253251;
    var search = 'https://api.themoviedb.org/3/' + queryParam + '/' + id + '/watch/providers?api_key=' + API_KEY // + '&language=es';
    var response = UrlFetchApp.fetch(search);
    var json = response.getContentText();
    var data = JSON.parse(json);
    var provider;
    //sort by type
    if ("ES" in data["results"]) {
        if ("flatrate" in data["results"]["ES"]) {
            provider = (data["results"]["ES"]["flatrate"][0]["provider_name"]);
        } else if ("rent" in data["results"]["ES"]) {
            provider = (data["results"]["ES"]["rent"][0]["provider_name"]);
        } else if ("buy" in data["results"]["ES"]) {
            provider = (data["results"]["ES"]["buy"][0]["provider_name"]);
        } else if ("free" in data["results"]["ES"]) {
            provider = (data["results"]["ES"]["free"][0]["provider_name"]);
        }
    } else if ("US" in data["results"]) {
        if ("flatrate" in data["results"]["US"]) {
            provider = (data["results"]["US"]["flatrate"][0]["provider_name"]);
        } else if ("rent" in data["results"]["US"]) {
            provider = (data["results"]["US"]["rent"][0]["provider_name"]);
        } else if ("buy" in data["results"]["US"]) {
            provider = (data["results"]["US"]["buy"][0]["provider_name"]);
        } else if ("free" in data["results"]["US"]) {
            provider = (data["results"]["US"]["free"][0]["provider_name"]);
        }
    } else {
        provider = 'Not found';
    }
    //rename
    if (provider == 'Amazon Prime Video') {
        provider = 'Amazon';
    } else if (provider == 'Movistar Plus') {
        provider = 'Movistar';
    } else if (provider == 'Disney Plus') {
        provider = 'Disney';
    } else if (provider == 'HBO Max') {
        provider = 'HBO';
    } else if (provider == 'Apple iTunes') {
        provider = 'Apple';
    } else if (provider == 'Google Play Movies') {
        provider = 'Google';
    } else if (provider == 'Rakuten TV') {
        provider = 'Rakuten';
    }
    return provider;
}

function integerToHours(value) {
    var hours = Math.floor(value / 60);
    var minutes = value % 60;
    //return hours + ":" + minutes;  
    return hours + 'h ' + minutes + ' min';
}
