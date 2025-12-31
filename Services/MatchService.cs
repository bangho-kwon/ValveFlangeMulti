using System;
using System.Collections.Generic;
using System.Linq;
using ValveFlangeMulti.Models;

namespace ValveFlangeMulti.Services
{
    public sealed class MatchResult
    {
        public PmsRow Valve { get; set; }
        public PmsRow Flange { get; set; }
        public PmsRow Gasket { get; set; } // optional
    }

    public sealed class MatchService
    {
        public MatchResult Match(List<PmsRow> rows, string selectedItemName, string pipeClass, double pipeSize)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            if (string.IsNullOrWhiteSpace(selectedItemName)) throw new ArgumentException("Selected ItemName is empty.");
            if (string.IsNullOrWhiteSpace(pipeClass)) throw new ArgumentException("Pipe Class is empty.");

            bool InRange(PmsRow r) => r.MainFrom <= pipeSize && pipeSize <= r.MainTo;

            // Alt empty only (for now)
            IEnumerable<PmsRow> BaseCandidates(string itemType) =>
                rows.Where(r =>
                    string.Equals(r.ItemType, itemType, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(r.Class, pipeClass, StringComparison.OrdinalIgnoreCase) &&
                    InRange(r) &&
                    string.IsNullOrWhiteSpace(r.Alt));

            var valve = BaseCandidates("Valve")
                .Where(r => string.Equals(r.ItemName, selectedItemName, StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();

            if (valve == null)
                throw new InvalidOperationException($"Valve 설정 없음: ItemName={selectedItemName}, Class={pipeClass}, Size={pipeSize}");

            if (string.IsNullOrWhiteSpace(valve.FamilyName) || string.IsNullOrWhiteSpace(valve.TypeName))
                throw new InvalidOperationException($"Valve FamilyName/TypeName 비어있음 (Row {valve.RowIndex})");

            var result = new MatchResult { Valve = valve };

            bool isFL = string.Equals(valve.ConnectionType ?? "", "FL", StringComparison.OrdinalIgnoreCase);
            if (isFL)
            {
                var flange = BaseCandidates("Flange").FirstOrDefault();
                if (flange == null)
                    throw new InvalidOperationException($"ConnectionType=FL 이지만 Flange 설정이 없습니다: Class={pipeClass}, Size={pipeSize}");

                if (string.IsNullOrWhiteSpace(flange.FamilyName) || string.IsNullOrWhiteSpace(flange.TypeName))
                    throw new InvalidOperationException($"Flange FamilyName/TypeName 비어있음 (Row {flange.RowIndex})");

                result.Flange = flange;

                var gasket = BaseCandidates("Gasket").FirstOrDefault();
                if (gasket != null)
                {
                    if (string.IsNullOrWhiteSpace(gasket.FamilyName) || string.IsNullOrWhiteSpace(gasket.TypeName))
                        throw new InvalidOperationException($"Gasket FamilyName/TypeName 비어있음 (Row {gasket.RowIndex})");
                    result.Gasket = gasket;
                }
            }

            return result;
        }
    }
}
