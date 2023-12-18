# sf2-contents

The tool reads the contents of the provided `SF2` soundfont file and generates `MD` and `XLSX` files with unpacked 
contents. The `MD` file provides the overview of the contents of the `SF2` file and the `XLSX` file can be used for a
detail study of the contents.

Refer to the `doc/sfspec24.pdf` file for format reference. 

## Usage
```
python sf2-contents.py <file>.sf2
```

### Output
```
<file>.md
<file>.xslx
```

## Dependencies

The tool uses [openpyxl](https://pypi.org/project/openpyxl/) to generate `XLSX` files.

## Soundfonts

Thousands of free soundfonts can be found all over the Net. Example:
https://sites.google.com/site/soundfonts4u/
