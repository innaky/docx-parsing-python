# Parsing docx using python 2.7

Parsing a docx file with python 2.7

# Explanation

The docx file has hundreds of lines.
It has information (in variable format) that is separated by a line with
symbols "------".
This script extract all intermediate sections to the boundary lines and save
this section in a text file, the name of this file is Header (variable format)
with the time (ex. 2018-01-31) the rest is data

# Example

## Docx file

```
-----------------------
Header date:2018-01-31T00:00:00
Information
-----------------------
Header date:2018-08-23T00:00:00
I
n
f
or
m
a
t
i
o
n
-----------------------
```

## Output

```
filename: Header-Time.txt
Information
```