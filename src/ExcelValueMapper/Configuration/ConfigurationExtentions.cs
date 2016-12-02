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
  using System.Collections.Generic;

  // Helper methods that work on configuration values
  public static class ConfigurationExtentions {
    // Preparing for handling column ranges
    public static IEnumerable<int> GetColumns ( this string columnValue ) {
      var result = 0;
      if ( !int.TryParse( columnValue, out result ) ) yield break;
      yield return result;
    }
  }
}
