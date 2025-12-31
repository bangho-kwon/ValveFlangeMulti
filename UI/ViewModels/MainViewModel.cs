using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using ValveFlangeCore.Placement;
using ValveFlangeCore.Models;
using ValveFlangeMulti.Models;
using ValveFlangeMulti.Settings;
using ValveFlangeMulti.Services;
using ValveFlangeMulti.UI.Enums;

using CoreInsertMode = ValveFlangeCore.Models.InsertMode;
namespace ValveFlangeMulti.UI.ViewModels
{
    public sealed class MainViewModel : INotifyPropertyChanged
    {
        private const int MaxStackTraceLength = 200; // 로그에 표시할 최대 스택 트레이스 길이
        
        private readonly ExternalCommandData _commandData;
        private readonly PmsExcelLoader _loader = new PmsExcelLoader();
        private readonly MatchService _matcher = new MatchService();

        private UserSettings _userSettings;

        private List<PmsRow> _rowsCache = new List<PmsRow>();

        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<ValveButtonItemViewModel> ValveItems { get; } = new ObservableCollection<ValveButtonItemViewModel>();
        public ICollectionView FilteredValveItems { get; private set; }

        private string _valveSearchText = "";
        public string ValveSearchText
        {
            get => _valveSearchText;
            set { _valveSearchText = value ?? ""; OnPropertyChanged(nameof(ValveSearchText)); FilteredValveItems.Refresh(); }
        }

        private ValveButtonItemViewModel _selectedValveItem;
        public ValveButtonItemViewModel SelectedValveItem
        {
            get => _selectedValveItem;
            set
            {
                _selectedValveItem = value;
                OnPropertyChanged(nameof(SelectedValveItem));
                OnPropertyChanged(nameof(SelectedValveDisplay));
                UiState = _selectedValveItem == null ? global::ValveFlangeMulti.UI.Enums.UiState.SettingLoaded : global::ValveFlangeMulti.UI.Enums.UiState.ValveSelected;
                RaiseCanExecute();
            }
        }

        public string SelectedValveDisplay => _selectedValveItem == null ? "Selected: -" : $"Selected: {_selectedValveItem.ItemName}";

        private UiState _uiState = global::ValveFlangeMulti.UI.Enums.UiState.NoSettingLoaded;
        public UiState UiState
        {
            get => _uiState;
            set { _uiState = value; OnPropertyChanged(nameof(UiState)); OnPropertyChanged(nameof(IsRunning)); RaiseCanExecute(); }
        }

        public bool IsRunning => UiState == global::ValveFlangeMulti.UI.Enums.UiState.Running;

        private SettingState _settingState = SettingState.NotLoaded;
        public SettingState SettingState
        {
            get => _settingState;
            set { _settingState = value; OnPropertyChanged(nameof(SettingState)); OnPropertyChanged(nameof(SettingStateText)); RaiseCanExecute(); }
        }

        public string SettingStateText => SettingState.ToString();

        private string _excelPathFull = "";
        public string ExcelPathFull
        {
            get => _excelPathFull;
            set { _excelPathFull = value ?? ""; OnPropertyChanged(nameof(ExcelPathFull)); OnPropertyChanged(nameof(ExcelPathDisplay)); }
        }

        public string ExcelPathDisplay => string.IsNullOrWhiteSpace(ExcelPathFull) ? "(not set)" : ExcelPathFull;

        private CoreInsertMode _insertMode = CoreInsertMode.Pick;
        public CoreInsertMode InsertMode
        {
            get => _insertMode;
            set { _insertMode = value; OnPropertyChanged(nameof(InsertMode)); ValidateOffset(); RaiseCanExecute(); }
        }

        private string _offsetMmText = "0";
        public string OffsetMmText
        {
            get => _offsetMmText;
            set { _offsetMmText = value ?? ""; OnPropertyChanged(nameof(OffsetMmText)); ValidateOffset(); RaiseCanExecute(); }
        }

        public double OffsetMmValue { get; private set; } = 0;
        public bool IsOffsetValid { get; private set; } = true;
        public string OffsetValidationMessage { get; private set; } = "";

        private string _statusMessage = "설정 파일을 불러오세요(PMS.xlsx).";
        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value ?? ""; OnPropertyChanged(nameof(StatusMessage)); }
        }

        public ObservableCollection<string> LogLines { get; } = new ObservableCollection<string>();

        public RelayCommand ImportRefreshCommand { get; }
        public RelayCommand RunSelectPipeCommand { get; }
        public RelayCommand SelectValveCommand { get; }

        public bool CanRun => SettingState == SettingState.Loaded && !IsRunning && SelectedValveItem != null && (InsertMode == CoreInsertMode.Pick || IsOffsetValid);

        public MainViewModel(ExternalCommandData commandData)
        {
            // Validate input
            if (commandData == null)
                throw new ArgumentNullException(nameof(commandData), "ExternalCommandData cannot be null");

            _commandData = commandData;

            // Load persisted settings (last Excel path) - best effort with defensive coding
            try
            {
                _userSettings = SettingsService.LoadGlobal();
                if (_userSettings != null && !string.IsNullOrWhiteSpace(_userSettings.LastExcelPath))
                    ExcelPathFull = _userSettings.LastExcelPath;
            }
            catch (Exception ex)
            {
                // Settings load failure shouldn't crash the app
                _userSettings = new UserSettings();
                LogLines.Add($"[WARN] Failed to load settings: {ex.Message}");
            }

            // Initialize filtered collection view with defensive null checks
            try
            {
                FilteredValveItems = CollectionViewSource.GetDefaultView(ValveItems);
                if (FilteredValveItems != null)
                {
                    FilteredValveItems.Filter = o =>
                    {
                        if (o is not ValveButtonItemViewModel v) return false;
                        if (string.IsNullOrWhiteSpace(ValveSearchText)) return true;
                        return v.ItemName.IndexOf(ValveSearchText, StringComparison.OrdinalIgnoreCase) >= 0;
                    };
                }
            }
            catch (Exception ex)
            {
                // Filter initialization failure - log but continue
                LogLines.Add($"[ERROR] Failed to initialize filter: {ex.Message}");
            }

            // Initialize commands with null checks
            ImportRefreshCommand = new RelayCommand(_ => ImportRefresh(), _ => !IsRunning);
            SelectValveCommand = new RelayCommand(p => SelectedValveItem = p as ValveButtonItemViewModel, _ => SettingState == SettingState.Loaded && !IsRunning);
            RunSelectPipeCommand = new RelayCommand(_ => RunSelectPipe(), _ => CanRun);

            ValidateOffset();
        }

        private void ImportRefresh()
        {
            try
            {
                var previousSelectedName = SelectedValveItem?.ItemName;

                SettingState = SettingState.Running;
                StatusMessage = "설정 파일을 불러오는 중…";

                // Import: let user choose PMS.xlsx (or any .xlsx)
                var fallbackPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PMS.xlsx");
                var initialPath = !string.IsNullOrWhiteSpace(ExcelPathFull) ? ExcelPathFull : fallbackPath;
                var initialDir = "";
                try
                {
                    initialDir = System.IO.Path.GetDirectoryName(initialPath) ?? "";
                }
                catch { initialDir = ""; }

                var dlg = new OpenFileDialog
                {
                    Title = "PMS.xlsx 선택",
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    FileName = System.IO.Path.GetFileName(initialPath),
                    CheckFileExists = true,
                    CheckPathExists = true,
                    InitialDirectory = System.IO.Directory.Exists(initialDir)
                        ? initialDir
                        : Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                };

                var ok = dlg.ShowDialog();
                if (ok != true)
                {
                    // User canceled: restore state gracefully
                    SettingState = string.IsNullOrWhiteSpace(ExcelPathFull) ? SettingState.NotLoaded : SettingState.Loaded;
                    UiState = (SettingState == SettingState.Loaded) ? global::ValveFlangeMulti.UI.Enums.UiState.SettingLoaded : global::ValveFlangeMulti.UI.Enums.UiState.NoSettingLoaded;
                    StatusMessage = "취소되었습니다.";
                    LogLines.Add("[INFO] (VFM-229) Import canceled.");
                    RaiseCanExecute();
                    return;
                }

                ExcelPathFull = dlg.FileName;

                // Persist last selected excel path
                _userSettings ??= new UserSettings();
                _userSettings.LastExcelPath = ExcelPathFull;
                SettingsService.SaveGlobal(_userSettings);

                _rowsCache = _loader.Load(ExcelPathFull);

                // Build valve buttons: all Valve ItemName (Alt empty only), group by ItemName, show classes
                var valves = _rowsCache
                    .Where(r => string.Equals(r.ItemType, "Valve", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(r.Alt))
                    .GroupBy(r => r.ItemName ?? "", StringComparer.OrdinalIgnoreCase)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g =>
                    {
                        var classes = g.Select(x => x.Class).Where(c => !string.IsNullOrWhiteSpace(c))
                            .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToList();
                        var classText = string.Join(", ", classes);
                        return new ValveButtonItemViewModel
                        {
                            ItemName = g.Key,
                            ClassListText = classText,
                            ClassListTooltip = classText
                        };
                    })
                    .OrderBy(v => v.ItemName)
                    .ToList();

                ValveItems.Clear();
                foreach (var v in valves) ValveItems.Add(v);
                FilteredValveItems.Refresh();

                // Import policy: try keep previous valve selection if still exists in new Excel.
                if (!string.IsNullOrWhiteSpace(previousSelectedName))
                {
                    var match = ValveItems.FirstOrDefault(x => string.Equals(x.ItemName, previousSelectedName, StringComparison.OrdinalIgnoreCase));
                    if (match != null)
                    {
                        SelectedValveItem = match;
                        StatusMessage = "설정 로드 완료. 파이프를 선택하세요.";
                        LogLines.Add("[INFO] (VFM-112) Selection kept after import.");
                    }
                    else
                    {
                        SelectedValveItem = null;
                        UiState = global::ValveFlangeMulti.UI.Enums.UiState.SettingLoaded;
                        StatusMessage = "설정 로드 완료. 밸브 종류를 선택하세요.";
                        LogLines.Add("[INFO] (VFM-113) Selection reset after import.");
                    }
                }
                else
                {
                    SelectedValveItem = null;
                    UiState = global::ValveFlangeMulti.UI.Enums.UiState.SettingLoaded;
                    StatusMessage = "설정 로드 완료. 밸브 종류를 선택하세요.";
                }

                SettingState = SettingState.Loaded;
                if (SelectedValveItem == null)
                    UiState = global::ValveFlangeMulti.UI.Enums.UiState.SettingLoaded;
                LogLines.Add("[INFO] (VFM-111) Settings loaded.");
                RaiseCanExecute();
            }
            catch (Exception ex)
            {
                SettingState = SettingState.Error;
                UiState = global::ValveFlangeMulti.UI.Enums.UiState.NoSettingLoaded;
                StatusMessage = $"설정 파일 로드 실패: {ex.Message}";
                LogLines.Add($"[ERROR] (VFM-121) {ex.Message}");
                RaiseCanExecute();
            }
        }

        private void RunSelectPipe()
        {
            UiState = global::ValveFlangeMulti.UI.Enums.UiState.Running;
            StatusMessage = "파이프를 선택하세요…";
            RaiseCanExecute();

            try
            {
                // Defensive null checks for Revit API objects
                if (_commandData?.Application?.ActiveUIDocument == null)
                    throw new InvalidOperationException("Active UI Document is not available");

                var uidoc = _commandData.Application.ActiveUIDocument;

                if (uidoc.Document == null)
                    throw new InvalidOperationException("Document is not available");

                if (uidoc.Selection == null)
                    throw new InvalidOperationException("Selection is not available");

                // Pick point on pipe (user just clicks the pipe)
                var pipeRef = uidoc.Selection.PickObject(ObjectType.PointOnElement, new PipeSelectionFilter(), "파이프를 선택하세요.");
                
                if (pipeRef == null)
                    throw new InvalidOperationException("Pipe selection returned null");

                var pickPoint = pipeRef.GlobalPoint;

                StatusMessage = "파이프 정보를 읽는 중… (SK_CLASS / Size)";
                var pipe = uidoc.Document.GetElement(pipeRef) as Autodesk.Revit.DB.Plumbing.Pipe;
                if (pipe == null) throw new InvalidOperationException("선택한 요소가 Pipe가 아닙니다.");

                string pipeClass = ReadStringParam(pipe, "SK_CLASS");
                if (string.IsNullOrWhiteSpace(pipeClass)) throw new InvalidOperationException("SK_CLASS를 읽을 수 없습니다.");

                double pipeSize = ReadPipeSizeMm(pipe); // in mm (simple heuristic)
                if (pipeSize <= 0) throw new InvalidOperationException("파이프 Size(Diameter)를 읽을 수 없습니다.");

                StatusMessage = "Excel 매칭 중… (Valve → Flange/Gasket)";
                
                if (SelectedValveItem == null)
                    throw new InvalidOperationException("밸브가 선택되지 않았습니다.");

                var match = _matcher.Match(_rowsCache, SelectedValveItem.ItemName, pipeClass, pipeSize);

                if (match?.Valve == null)
                    throw new InvalidOperationException("매칭된 밸브 정보가 없습니다.");

                // Build request
                var req = new PlacementRequest(_commandData, pipeRef, pickPoint)
                {
                    Mode = (ValveFlangeCore.Models.InsertMode)Enum.Parse(typeof(ValveFlangeCore.Models.InsertMode), InsertMode.ToString()),
                    OffsetMm = (InsertMode == CoreInsertMode.Pick) ? 0 : OffsetMmValue,
                    ValveFamilyName = match.Valve.FamilyName,
                    ValveTypeName = match.Valve.TypeName,
                    ConnectionType = match.Valve.ConnectionType ?? ""
                };

                if (string.Equals(req.ConnectionType, "FL", StringComparison.OrdinalIgnoreCase))
                {
                    if (match.Flange != null)
                    {
                        req.FlangeFamilyName = match.Flange.FamilyName;
                        req.FlangeTypeName = match.Flange.TypeName;
                    }

                    if (match.Gasket != null)
                    {
                        req.GasketFamilyName = match.Gasket.FamilyName;
                        req.GasketTypeName = match.Gasket.TypeName;
                    }
                    else
                    {
                        LogLines.Add("[WARN] (VFM-331) Gasket 설정 없음 → 생략");
                    }
                }

                StatusMessage = "배치 중… (파이프 분할 및 연결)";
                string msg = null;
                var coreResult = ValveFlangePlacementService.Execute(_commandData, req, ref msg);

                if (!coreResult.Success)
                    throw new InvalidOperationException(coreResult.Message);

                StatusMessage = "배치 완료 ✅";
                LogLines.Add("[INFO] (VFM-511) Completed.");
                UiState = global::ValveFlangeMulti.UI.Enums.UiState.ValveSelected;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                StatusMessage = "취소되었습니다.";
                LogLines.Add("[INFO] (VFM-520) User canceled operation.");
                UiState = global::ValveFlangeMulti.UI.Enums.UiState.ValveSelected;
            }
            catch (Exception ex)
            {
                StatusMessage = $"오류: {ex.Message}";
                LogLines.Add($"[ERROR] (VFM-521) {ex.GetType().Name}: {ex.Message}");
                
                // Log stack trace for debugging
                if (!string.IsNullOrWhiteSpace(ex.StackTrace))
                {
                    LogLines.Add($"[DEBUG] Stack trace: {ex.StackTrace.Substring(0, Math.Min(MaxStackTraceLength, ex.StackTrace.Length))}");
                }
                
                UiState = global::ValveFlangeMulti.UI.Enums.UiState.ValveSelected;
            }
            finally
            {
                RaiseCanExecute();
            }
        }

        private void ValidateOffset()
        {
            if (InsertMode == CoreInsertMode.Pick)
            {
                IsOffsetValid = true;
                OffsetValidationMessage = "";
                OffsetMmValue = 0;
                OnPropertyChanged(nameof(IsOffsetValid));
                OnPropertyChanged(nameof(OffsetValidationMessage));
                return;
            }

            if (string.IsNullOrWhiteSpace(OffsetMmText))
            {
                IsOffsetValid = false;
                OffsetValidationMessage = "Offset 입력이 필요합니다.";
                OnPropertyChanged(nameof(IsOffsetValid));
                OnPropertyChanged(nameof(OffsetValidationMessage));
                return;
            }

            if (!double.TryParse(OffsetMmText, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) &&
                !double.TryParse(OffsetMmText, NumberStyles.Any, CultureInfo.CurrentCulture, out v))
            {
                IsOffsetValid = false;
                OffsetValidationMessage = "Offset은 숫자여야 합니다.";
                OnPropertyChanged(nameof(IsOffsetValid));
                OnPropertyChanged(nameof(OffsetValidationMessage));
                return;
            }

            if (v < 0 || v > 5000)
            {
                IsOffsetValid = false;
                OffsetValidationMessage = "Offset 범위: 0~5000mm";
                OnPropertyChanged(nameof(IsOffsetValid));
                OnPropertyChanged(nameof(OffsetValidationMessage));
                return;
            }

            IsOffsetValid = true;
            OffsetValidationMessage = "";
            OffsetMmValue = v;
            OnPropertyChanged(nameof(IsOffsetValid));
            OnPropertyChanged(nameof(OffsetValidationMessage));
        }

        private void RaiseCanExecute()
        {
            ImportRefreshCommand.RaiseCanExecuteChanged();
            RunSelectPipeCommand.RaiseCanExecuteChanged();
            SelectValveCommand.RaiseCanExecuteChanged();
            OnPropertyChanged(nameof(CanRun));
        }

        private void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private static string ReadStringParam(Element e, string paramName)
        {
            var p = e.LookupParameter(paramName);
            return p?.AsString() ?? "";
        }

        private static double ReadPipeSizeMm(Autodesk.Revit.DB.Plumbing.Pipe pipe)
        {
            // Tries common built-in params first, falls back to connector diameter.
            var p = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            if (p != null && p.StorageType == StorageType.Double)
            {
                var ft = p.AsDouble();
                return UnitUtils.ConvertFromInternalUnits(ft, UnitTypeId.Millimeters);
            }

            // Fallback: connector
            var conns = pipe.ConnectorManager?.Connectors;
            if (conns != null)
            {
                foreach (Connector c in conns)
                {
                    var r = c.Radius;
                    if (r > 0)
                        return UnitUtils.ConvertFromInternalUnits(r * 2.0, UnitTypeId.Millimeters);
                }
            }

            return 0;
        }

        private sealed class PipeSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is Autodesk.Revit.DB.Plumbing.Pipe;
            public bool AllowReference(Reference reference, XYZ position) => true;
        }
    }

    public enum SettingState { NotLoaded, Loaded, Error, Running }
}
