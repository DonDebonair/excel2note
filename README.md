# Excel2Note

Excel2note is a utility that converts Excel sheets to a collection of Evernote notes. It takes an existing Excel file as
input together with the name of the output file you wish to create. The output file is of the _.enex_ format and can be
imported directly into Evernote.

## Installation

You can install excel2note by cloning it and running the provided setup file:

```
git clone https://github.com/DandyDev/excel2note.git
python setup.py install
```

## Usage

Excel2Note can be called as follows: `excel2note [-h] [-n <row nr>] [-t <column nr>] excel_file enex_file`

If your column names are not on the first row (1-based), than you have to provide the row number using `-n` or `--names`.
Excel2note will use the text in the first column of each row as title for the note it creates. You can override this
by supplying another column number using `-t` or `--title`.

The resulting output file can be imported into Evernote. All notes will be put into a single notebook.

## Contact

If you have an problem, please create an issue here on GitHub or email me at: [debie.daan@gmail.com](mailto:debie.daan@gmail.com)