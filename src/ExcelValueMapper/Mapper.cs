//*******************************************************************************************
// ExcelValueMapper - Mapper.cs
//*******************************************************************************************

// Copyright 2016 Sascha Andres

// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at

//     http://www.apache.org/licenses/LICENSE-2.0

// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

namespace ExcelValueMapper {
  using Configuration;
  using OfficeOpenXml;
  using System;
  using System.Collections.Generic;
  using System.Linq;

  public class Mapper {
    private const int MINIMUM_INDEX_EXCEL_ADDRESS = 1;

    private readonly ExcelPackage _Package;
    private Config _Configuration;

    public Config Configuration {
      get { return _Configuration; }
      internal set { _Configuration = value; }
    }

    #region Contructors
    // Default constructor, usable to work with <see cref="Configuration"/>
    public Mapper () { }

    // Constructor to be used when actually running a mapping
    public Mapper ( ExcelPackage package, Config config ) {
      _Package = package;
      Configuration = config;
    }
    #endregion

    #region Validation
    // IsValid returns true if the ValueTranslator is setup technically correct
    public bool IsValid () {
      return Validate().Count() == 0;
    }

    // Validate returns errors in setup
    // Checks for null excel document, no replacement map and 
    // that no values are configured for two replacements
    public IEnumerable<string> Validate () {
      // 28-Nov-2016 - SMA - TODO: Keys for Maps and MapForColumn should match
      if ( null == _Package ) {
        yield return "ERROR: No Excel document provided";
      } else {
        if ( null == Configuration || null == Configuration.Maps || 0 == Configuration.Maps.Count ) {
          yield return "ERROR: Configuration is not valid";
        } else {
          foreach ( var map in Configuration.Maps ) {
            foreach ( var currentKey in map.Value.Keys ) {
              foreach ( var replacementValue in map.Value[currentKey] ) {
                foreach ( var compareKey in map.Value.Keys ) {
                  if ( compareKey == currentKey ) continue;
                  foreach ( var compareReplacementValue in map.Value[compareKey] ) {
                    if ( compareReplacementValue == replacementValue )
                      yield return $"WARNING: {replacementValue} ({currentKey}) is also mapped to {compareKey} in {map.Key}";
                  }
                }
              }
            }
          }
        }
      }
    }
    #endregion

    #region Value calculation, dump of mapping
    private string calculateCellValue ( string cellValue, IDictionary<string, string> map ) {
      if ( map.ContainsKey( cellValue ) ) {
        return map[cellValue];
      }
      switch ( Configuration.UnknownValueHandling ) {
        case enum_UnknownValueHandling.Default:
          return Configuration.Default;
        case enum_UnknownValueHandling.Leave:
          return cellValue;
        case enum_UnknownValueHandling.Empty:
          return string.Empty;
      }
      throw new Exception( $"ERROR: Cell value could not be calculated ({cellValue})" );
    }

    private void writeValueList () {
      if ( Configuration.WriteMappingToWorksheet <= 0 ) return;
      var workingSheet = Configuration.WriteMappingToWorksheet;
      if ( _Package.Workbook.Worksheets.Count >= workingSheet ) {
        var rowIndex = MINIMUM_INDEX_EXCEL_ADDRESS;
        foreach ( string mapKey in Configuration.Maps.Keys ) {
          var columnIndex = MINIMUM_INDEX_EXCEL_ADDRESS;
          _Package.Workbook.Worksheets[workingSheet].Cells[rowIndex, columnIndex].Value = mapKey;
          columnIndex++;
          foreach (var mapping in Configuration.Maps[mapKey] ) {
            _Package.Workbook.Worksheets[workingSheet].Cells[rowIndex, columnIndex].Value = $"{mapping.Key} := {mapping.Value}";
            columnIndex++;
          }
          rowIndex++;
        }
      }
    }
    #endregion

    // Translate iterates through the cells of worksheet one and calculates the mapping value for the cell
    public void Translate () {
      if ( Validate().Any( _ => _.StartsWith( "ERROR" ) ) ) {
        throw new Exception( "ERROR: Configuration contains errors" );
      }

      var workingWorksheet = Math.Max( MINIMUM_INDEX_EXCEL_ADDRESS, Configuration.MapWorksheet );

      foreach ( string mapKey in Configuration.Maps.Keys ) {
        if ( !Configuration.MapForColumn.ContainsKey( mapKey ) ) continue;
        var map = Configuration.Maps[mapKey];
        var columns = Configuration.MapForColumn[mapKey];
        foreach ( string columnList in columns ) {
          foreach ( int columnIndex in columnList.GetColumns() ) {
            var rowIndex = Configuration.StartAtLine;
            while ( cellHasData( workingWorksheet, rowIndex, columnIndex ) ) {
              _Package.Workbook.Worksheets[workingWorksheet].Cells[rowIndex, columnIndex].Value = calculateCellValue( getCellValueAsStringOrEmpty( workingWorksheet, rowIndex, columnIndex ), map );
              rowIndex++;
            }
          }
        }
      }

      writeValueList();
    }

    #region Excel helper methods
    private string getCellValueAsStringOrEmpty ( int worksheet, int rowIndex, int columnIndex ) {
      return (_Package.Workbook.Worksheets[worksheet].Cells[rowIndex, columnIndex].Value ?? string.Empty).ToString();
    }

    private bool cellHasData ( int worksheet, int rowIndex, int columnIndex ) {
      return null != _Package.Workbook.Worksheets[worksheet].Cells[rowIndex, columnIndex].Value;
    }
    #endregion
  }
}