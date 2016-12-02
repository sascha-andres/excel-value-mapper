//*******************************************************************************************
// ExcelValueMapper - Config.cs
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
  using System;
  using System.Collections.Generic;
  using System.Linq;
  using System.Text;
  using System.Threading.Tasks;

  // Configuration for a mapping
  public class Config {
    /// <summary>
    /// Initializes a new instance of the <see cref="Config"/> class.
    /// </summary>
    public Config () {
      WriteMappingToWorksheet = 0;
      UnknownValueHandling = enum_UnknownValueHandling.Leave;
      Default = String.Empty;
      Maps = new Dictionary<string, Dictionary<string, string>>();
      MapForColumn = new Dictionary<string, List<string>>();
      StartAtLine = 2;
      
    }
    
    // If set to a value not equal to 0 a list of mappings will be dumped
    // to the worksheet with that index
    public int WriteMappingToWorksheet { get; set; }

    // Start reading values at line #s
    public int StartAtLine { get; set; }
    
    // How to handle unknown values
    public enum_UnknownValueHandling UnknownValueHandling { get; set; }
    
    // Default value, defaulting to string.Empty
    public string Default { get; set; }
    
    // Lists of maps
    public Dictionary<string, Dictionary<string, string>> Maps { get; set; }
    
    // Matches maps (the keys are the same as in <see cref="Maps"/>) to columns
    // Currently only column numbers, later column ranges
    public Dictionary<string, List<string>> MapForColumn { get; set; }
    
    // What worksheet to map
    public int MapWorksheet { get; set; }
  }
}
