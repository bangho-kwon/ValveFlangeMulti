using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ValveFlangeMulti.Models;

namespace ValveFlangeMulti.Services
{
    public sealed class PmsExcelLoader
    {
        // Expected columns (1-based): A=Class, C=MainFrom, D=MainTo, G=Alt, H=ItemType, I=ItemName, L=ConnectionType, M=FamilyName, N=TypeName
        public List<PmsRow> Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) 
                throw new ArgumentException("Excel path is empty.", nameof(path));

            if (!System.IO.File.Exists(path))
                throw new System.IO.FileNotFoundException($"Excel file not found: {path}");

            try
            {
                using var wb = new XLWorkbook(path);
                
                if (wb.Worksheets == null || wb.Worksheets.Count == 0)
                    throw new InvalidOperationException("Workbook has no worksheets.");

                var ws = wb.Worksheets.First();
                if (ws == null)
                    throw new InvalidOperationException("Failed to get first worksheet.");

                // Find first used row (assume header at row 1)
                var used = ws.RangeUsed();
                if (used == null) throw new InvalidOperationException("Worksheet is empty.");

                int firstDataRow = used.RangeAddress.FirstAddress.RowNumber + 1; // after header
                int lastRow = used.RangeAddress.LastAddress.RowNumber;

                var rows = new List<PmsRow>();
                for (int r = firstDataRow; r <= lastRow; r++)
                {
                    try
                    {
                        string cls = ws.Cell(r, 1).GetString().Trim(); // A
                        if (string.IsNullOrWhiteSpace(cls)) continue; // skip empty lines

                        double mainFrom = ReadDouble(ws.Cell(r, 3)); // C
                        double mainTo = ReadDouble(ws.Cell(r, 4));   // D

                        var row = new PmsRow
                        {
                            RowIndex = r,
                            Class = cls,
                            MainFrom = mainFrom,
                            MainTo = mainTo,
                            Alt = ws.Cell(r, 7).GetString().Trim(),
                            ItemType = ws.Cell(r, 8).GetString().Trim(),
                            ItemName = ws.Cell(r, 9).GetString().Trim(),
                            ConnectionType = ws.Cell(r, 12).GetString().Trim(),
                            FamilyName = ws.Cell(r, 13).GetString().Trim(),
                            TypeName = ws.Cell(r, 14).GetString().Trim(),
                        };

                        rows.Add(row);
                    }
                    catch (Exception ex)
                    {
                        // Log row-level errors but continue processing
                        System.Diagnostics.Debug.WriteLine($"Error reading row {r}: {ex.Message}");
                    }
                }

                return rows;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to load Excel file: {ex.Message}", ex);
            }
        }

        private static double ReadDouble(IXLCell cell)
        {
            try
            {
                if (cell == null) return 0.0;
                if (cell.DataType == XLDataType.Number) return cell.GetDouble();
                
                var s = cell.GetString().Trim();
                if (string.IsNullOrWhiteSpace(s)) return 0.0;
                
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return v;
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out v)) return v;
            }
            catch
            {
                // Return 0 on any parse error
            }
            return 0.0;
        }
    }
}
