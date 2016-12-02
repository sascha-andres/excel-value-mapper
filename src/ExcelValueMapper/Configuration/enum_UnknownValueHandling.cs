//*******************************************************************************************
// ExcelValueMapper - enum_UnknownValueHandling.cs
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
  // How to map not configured values
  public enum enum_UnknownValueHandling {
    // set to default
    Default,
    // leave them
    Leave,
    // Set empty
    Empty
  }
}
