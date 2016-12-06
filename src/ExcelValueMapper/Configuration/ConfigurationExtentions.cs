//*******************************************************************************************
// ExcelValueMapper - ConfigurationExtentions.cs
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

namespace ExcelValueMapper.Configuration {
  using System.Text.RegularExpressions;
  using System.Collections.Generic;

  // Helper methods that work on configuration values
  public static class ConfigurationExtentions {
    private const string RANGE_EXPRESSION = "^(?<from>[0-9]*)-(?<to>[0-9]*)$";
    private static Regex _Regex = new Regex( RANGE_EXPRESSION );

    // Preparing for handling column ranges
    public static IEnumerable<int> GetColumns ( this string columnValue ) {
      if ( _Regex.IsMatch( columnValue ) ) {
        var matchedResult = _Regex.Match( columnValue );
        var from = 0;
        var to = 0;
        if ( !int.TryParse( matchedResult.Groups["from"].Value, out from ) ) {
          System.Console.WriteLine( $"WARNING: {columnValue} seems to be a range but {matchedResult.Groups["from"].Value} is not parsable" );
          yield break;
        }
        if ( !int.TryParse( matchedResult.Groups["to"].Value, out to ) ) {
          System.Console.WriteLine( $"WARNING: {columnValue} seems to be a range but {matchedResult.Groups["to"].Value} is not parsable" );
          yield break;
        }
        if (to <= from) {
          System.Console.WriteLine( $"WARNING: {to} is <= {from}, dropping" );
          yield break;
        }
        for ( int i = from; i <= to; i++ ) {
          yield return i;
        }
      } else {
        var result = 0;
        if ( !int.TryParse( columnValue, out result ) ) {
          System.Console.WriteLine( $"WARNING: {columnValue} is non numeric and not a range" );
          yield break;
        }
        yield return result;
      }
    }
  }
}
