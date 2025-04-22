namespace XlBlocks.AddIn.Dna;

using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using NLog;
using XlBlocks.AddIn.Gui;
using XlBlocks.AddIn.Properties;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private static readonly Logger _logger = LogManager.GetLogger(typeof(RibbonController).FullName);

    private IRibbonUI? _ribbonUI;
    public override string GetCustomUI(string RibbonID)
    {
        return Resources.ribbon_xml;
    }

    public void Ribbon_Load(IRibbonUI ribbon)
    {
        _logger.Debug("loading ribbon UI");
        _ribbonUI = ribbon;
        GuiHelper.RegisterRibbon(_ribbonUI);
    }

    public object? Ribbon_GetImage(IRibbonControl control)
    {
        return control.Id switch
        {
            "button_erroroutput" => Ribbon_LoadImage($"error_output_{(XlBlocksAddIn.PrintExceptions ? "on" : "off")}"),
            _ => null,
        };
    }

    public object? Ribbon_LoadImage(string imageId)
    {
        try
        {
            var imageObj = Resources.ResourceManager.GetObject(imageId);
            if (imageObj == null)
                return null;

            if (imageObj is Image image)
            {
                return image;
            }
            else if (imageObj is byte[] imageBytes)
            {
                using var imageStream = new MemoryStream(imageBytes);
                return Image.FromStream(imageStream);
            }
            return null;
        }
        catch (Exception ex)
        {
            _logger.Error(ex);
            return null;
        }
    }

    public string Ribbon_GetLabel(IRibbonControl control)
    {
        return control.Id switch
        {
            "button_erroroutput" => $"Error Output {(XlBlocksAddIn.PrintExceptions ? "On" : "Off")}",
            _ => "",
        };
    }

    public void Ribbon_OnButtonPressed(IRibbonControl control)
    {
        switch (control.Id)
        {
            case "button_about":
                GuiHelper.ShowAboutForm();
                break;
            case "button_viewcachedobject":
                _logger.Debug("Show cached object");
                break;
            case "button_viewcache":
                GuiHelper.ShowCacheViewerForm();
                break;
            case "button_clearcache":
                GuiHelper.ShowCacheClearVerify();
                break;
            case "button_erroroutput":
                XlBlocksAddIn.PrintExceptions = !XlBlocksAddIn.PrintExceptions;
                break;
            case "button_errorlog":
                GuiHelper.ShowErrorLogViewerForm();
                break;
        }
    }
}
