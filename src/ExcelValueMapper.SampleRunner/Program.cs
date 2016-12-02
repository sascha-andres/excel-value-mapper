//*******************************************************************************************
// ExcelValueMapper.SampleRunner - Program.cs
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

using System.Diagnostics.Contracts;
namespace ExcelValueMapper.SampleRunner {
  using System.IO;
  using OfficeOpenXml;
  using ExcelValueMapper.Configuration;
  using System;
  using System.Collections.Generic;
  using System.Linq;
  using System.Text;
  using System.Threading.Tasks;
  using static System.Console;
  class Program {
    static int Main ( string[] args ) {
      Contract.Assert( args.Length == 3, "Usage: executable <config> <excel> <newexcel>" );
      Contract.Assert( File.Exists( args[0] ), "Configuration file does not exist" );
      Contract.Assert( File.Exists( args[1] ), "Source Excel workbook does not exist" );
      Contract.Assert( !File.Exists( args[2] ), "Destination Excel workbook does exist" );

      using ( var pkg = new ExcelPackage( new FileInfo( args[1] ) ) ) {
        var mapper = new Mapper( pkg, new Config() );
        mapper.LoadConfiguration( args[0] );
        var warningsAndErros = mapper.Validate().ToList();
        if ( warningsAndErros.Count > 0 ) {
          foreach ( var item in warningsAndErros ) {
            WriteLine( item );
          }
          if ( warningsAndErros.Any( _ => _.StartsWith( "ERROR" ) ) ) return -2;
        }
        mapper.Translate();
        pkg.SaveAs( new FileInfo( args[2] ) );
      }
      return 0;
    }
  }
}
