//*******************************************************************************************
// ExcelValueMapper - MapperExtentions.cs
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
  using System.IO;
  using YamlDotNet.Serialization;

  // Helper methods working on the Mapper type
  public static class MapperExtentions {

    // Save configuration from instance
    public static void SaveConfiguration ( this Mapper value, string path ) {
      var serializer = new Serializer();
      File.Delete( path );
      File.WriteAllText( path, serializer.Serialize( value.Configuration ) );
    }

    // Load configuration to instance
    public static void LoadConfiguration ( this Mapper value, string path ) {
      if ( !File.Exists( path ) ) throw new Exception( "ERROR: Configuration file does not exist" );
      var deserializer = new Deserializer();
      value.Configuration = deserializer.Deserialize<Config>( File.ReadAllText( path ) );
    }

    // Create an empty configuration
    public static void CreateEmptyConfiguration ( this Mapper value ) {
      value.Configuration = new Config();
    }
  }
}
