# Excel value mapper

## Purpose

Change values in an Excel file eg to have tools like SPSS working on numerical instead of textual values.

## Usage

This is a library to use in yout programs. Two demonstration programs are included:

* evm ( Project `ExcelValueMapper.SampleRunner` )
* sampleconfig ( `Project ExcelValueMapper.SampleConfigurationGenerator` )

### evm

`evm` stands for ExcelValueMapper and is a sample runner. You have to call it with three parameters ( in that order ):

1. Config file
2. Excel file input
3. Excel file output

The latter Excel file must not exist.

### sampleconfig

Just creates a sample configuration ( like the one in the `sample` directory )

## Sample

### Configuration

    WriteMappingToWorksheet: 2
    StartAtLine: 2
    UnknownValueHandling: Leave
    Default: DefaultValue
    Maps:
      map1:
        value 1: 0
        value 2: 1
      map2:
        value 1: 1
        value 2: 0
    MapForColumn:
      map1:
      - 1
      - 2
      map2:
      - 3
      - 4
      - 5
    MapWorksheet: 1

|Configuration value|Description|
|---|---|
|WriteMappingToWorksheet|Write mapping summary to worksheet with this index, based 1|
|StartAtLine|To skip headlines set it to the line to start replace at, based 1|
|UnknownValueHandling|Leave: do not change|
||Default: use provided default|
||Empty: Write an empty cell|
|Maps|A list of named mappings|
|MapForColumn|Provide list of column indices for each map|
|MapWorksheet|Do mapping in this worksheet, based 1|

### Sample input

|map1-1|map1-2|map2-1|map2-2|map2-3|map1-12|map1-22|map2-12|map2-22|map2-32|
|---|---|---|---|---|---|---|---|---|---|
|value 1|value 2|value 1|value 1|value 1|value 1|value 2|value 1|value 1|value 1|
|value 2|value 1|value 2|value 2|value 2|value 2|value 1|value 2|value 2|value 2|
|unknown|unknown|unknown|unknown|unknown|unknown|unknown|unknown|unknown|unknown|

### Sample output

|map1-1|map1-2|map2-1|map2-2|map2-3|map1-12|map1-22|map2-12|map2-22|map2-32|
|---|---|---|---|---|---|---|---|---|---|
|0|1|1|1|1|value 1|value 2|value 1|value 1|value 1|
|1|0|0|0|0|value 2|value 1|value 2|value 2|value 2|
|unknown|unknown|unknown|unknown|unknown|unknown|unknown|unknown|unknown|unknown|

### Sample mapping output

||||
|---|---|---|---|
|map1|value 1 := 0|value 2 := 1|
|map2|value 1 := 1|value 2 := 0|

## History

|Version|Contributors|Changes|
|---|---|---|
|20161202|S. Andres|Initial code|