import xlsx from 'xlsx';
import fs from 'fs';
import config from '../config';

const SEPARATOR = "###";

class Excel {
  constructor() {
    this.book = xlsx.readFile( config.excel.InputExcelFilePath );
    this.sheet = this.book.Sheets[config.excel.SheetName];
  }

  exec() {
    const excelData = this.getExcelData();
    this.createJsonFile( excelData );
  }

  getExcelData() {
    const range = this.sheet['!ref'];
    console.log( 'range:' + range );
    const rangeVal = xlsx.utils.decode_range( range );
    const maxColumnNumber = this.getMaxColumnNumber();

    let map = new Map();
    let text = '';
    for ( let row = rangeVal.s.r; row <= rangeVal.e.r; row++ ) {
      //見出し行はとばす
      if ( row <= config.excel.HeadRowNumber ) {
        continue;
      }
      for ( let column = rangeVal.s.c; column <= rangeVal.e.c; column++ ) {
        let adr = xlsx.utils.encode_cell( { c: column, r: row } );
        let cell = this.sheet[adr];
        this.setMap( column, map, cell, adr )
      }
      text += `{"index":{}}\n`;
      text += `{"text":"`;
      Object.keys( config.excel.Columns ).forEach( function ( key ) {
        text += `${map.get( this[key] )}`;
        if ( maxColumnNumber !== key ) {
          text += `${SEPARATOR}`;
        }
      }, config.excel.Columns );
      text += `"}\n`;
      map.clear();
    }
    return text;
  }

  getMaxColumnNumber() {
    for ( var i in config.excel.Columns ) {
    }
    return i;
  }

  setMap( column, map, cell, adr ) {
    if ( column in config.excel.Columns ) {
      if ( cell && cell.v ) {
        map.set( config.excel.Columns[column], cell.v );
      } else {
        map.set( config.excel.Columns[column], '' );
        console.log( '空白のセル:' + JSON.stringify( adr ) );
      }
    }
  }

  createJsonFile( excelData ) {
    fs.writeFile( config.excel.OutputJsonFilePath, excelData, function ( err ) {
      if ( err ) {
        console.log( err );
      } else {
        console.log( 'success' );
      }
    } );
  }
}

export default Excel