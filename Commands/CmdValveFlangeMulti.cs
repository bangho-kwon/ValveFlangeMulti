using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using System;
using ValveFlangeMulti.UI;

namespace ValveFlangeMulti.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CmdValveFlangeMulti : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            try
            {
                // Validate command data
                if (commandData == null)
                {
                    message = "ExternalCommandData is null";
                    return Result.Failed;
                }

                if (commandData.Application == null)
                {
                    message = "UIApplication is null";
                    return Result.Failed;
                }

                // Create and show window
                var win = new ValveFlangeMultiWindow(commandData);
                if (win == null)
                {
                    message = "Failed to create window";
                    return Result.Failed;
                }

                win.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"ValveFlangeMulti command failed: {ex.Message}";
                
                // Log to TaskDialog for user visibility
                try
                {
                    TaskDialog.Show("Error", $"ValveFlangeMulti encountered an error:\n\n{ex.Message}\n\nDetails:\n{ex.GetType().Name}");
                }
                catch
                {
                    // Ignore errors in error display
                }

                return Result.Failed;
            }
        }
    }
}
