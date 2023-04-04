# ExcelUtilities
Useful excel macros for all kinds of stuff.

This repository was created to keep me from writing the same code again and again.

# Type conversion

|Name|Description|
|:----|----|
|[CtoLng](arrays.bas)|conversiont to Long type regardless of used decimal separator.|
|[fConv](arrays.bas)|conversion of decimal number into fixed length number of variable base. Can be also used as a counter with customizable scales. For more details check the function description in the module.|
|[DUMP](arrays.bas)|function serializes (almost) any type of variable. It was originaly meant to serialize a dictionary, but because a dictionary item can be of any type, also this function needs to dump any type also including arrays reagardless of number of dimensions.|