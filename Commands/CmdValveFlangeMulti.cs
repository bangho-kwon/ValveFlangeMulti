using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using ValveFlangeMulti.UI;

namespace ValveFlangeMulti.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CmdValveFlangeMulti : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            var win = new ValveFlangeMultiWindow(commandData);
            win.ShowDialog();
            return Result.Succeeded;
        }
    }
}
