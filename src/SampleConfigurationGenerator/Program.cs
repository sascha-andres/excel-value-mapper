//*******************************************************************************************
// SampleConfigurationGenerator - Program.cs
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

namespace ExcelValueMapper.SampleConfigurationGenerator {
  using Configuration;
  using System.Collections.Generic;

  class Program {
    static void Main ( string[] args ) {
      var evm = new Mapper();
      evm.CreateEmptyConfiguration();
      evm.Configuration.Default = "DefaultValue";
      evm.Configuration.Maps["map1"] = new Dictionary<string, string>();
      evm.Configuration.Maps["map1"]["value 1"] = "0";
      evm.Configuration.Maps["map1"]["value 2"] = "1";
      evm.Configuration.Maps["map2"] = new Dictionary<string, string>();
      evm.Configuration.Maps["map2"]["value 1"] = "1";
      evm.Configuration.Maps["map2"]["value 2"] = "0";
      evm.Configuration.MapForColumn["map1"] = new List<string>();
      evm.Configuration.MapForColumn["map2"] = new List<string>();
      evm.Configuration.MapForColumn["map1"].Add( "1" );
      evm.Configuration.MapForColumn["map1"].Add( "2" );
      evm.Configuration.MapForColumn["map2"].Add( "3" );
      evm.Configuration.MapForColumn["map2"].Add( "4" );
      evm.Configuration.MapForColumn["map2"].Add( "5" );
      evm.Configuration.WriteMappingToWorksheet = 2;
      evm.Configuration.UnknownValueHandling = enum_UnknownValueHandling.Leave;
      evm.Configuration.MapWorksheet = 1;
      evm.Configuration.WriteMappingToWorksheet = 2;
      evm.SaveConfiguration( "config.yaml" );
    }
  }
}
