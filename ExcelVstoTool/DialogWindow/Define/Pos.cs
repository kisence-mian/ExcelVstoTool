using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

struct Pos
{
    public int col;
    public int row;

    public Pos(Range range)
    {
        col = range.Column;
        row = range.Row;
    }
}