<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="JETNET_RESEARCH_ASSISTANT._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
  <title>Research Assistant.</title>
  
  
  <script language="javascript" type="text/javascript">
    var bDontClose = false;
 

    function openReportWindow(reportWindowPath, sReportID) {

      //alert(" show report : " + bShowReport + " report path : " + reportWindowPath + " report number : " + sReportID);

      var rightNow = new Date();
      var reportWindowName = "AdminReport" + sReportID + "Window";
      reportWindowName += rightNow.getTime();

      var reportWindowOptions = "scrollbars=yes,menubar=yes,height=800,width=1050,resizable=yes,toolbar=no,location=no,status=no";

      if (reportWindowPath != "") {
        var Place = window.open(reportWindowPath, reportWindowName, reportWindowOptions);
      }

      return true;
    }
         
  </script>

  <script src="https://use.fontawesome.com/52d48867c2.js"></script>

  <link rel="stylesheet" href="/css/skeleton_grid.css" type="text/css" />
  <link rel="stylesheet" href="/css/theme.css" type="text/css" />
</head>
<body>
  <form id="form1" runat="server">
  <div class="FixedHeaderBar">
  </div>
  <div class="container">
    <div class="sixteen columns headerHeight">
      <div class="one-third column logo">
        <img alt="Research Assistant" src="pictures/ResearchLogo.png" />
      </div>
    </div>
    <div class="sixteen columns main">
      <asp:Label runat="server" ID="text_label"></asp:Label>
      <asp:Label runat="server" ID="integrity_label"></asp:Label>
    </div>
  </div>
  </form>
</body>
</html>
