# SharePoint
Usefull SharePoint JavaScript solutions

<div>
<h1>Ribbon_UserCustomActions</h1>
</div>

Adds Ribbon/User Custom Action to send selected Documents as attachments.
Works only in IE(requires ActiveX) with Outlook.
 
<h2>Files<h2>
<ul>
    <li><i>SendDocument.js</i> - Opens Outlook to send email with Attachments</li>
    <li><i>send-docs-customaction.js</i> - User Custom Actions logic</li>
    <li><i>send-docs-ribbon.js</i> - Ribbon operations</li>
    <li><i>send-docs-main.js</i> -  Main function is here</li>
</ul>
 
 
<div>
<h2>Usage</h2>
To use just oput links to specified scripts on your page: e.g. put in master page markup:
<!--SPM:<SharePoint:ScriptLink language="javascript" name="~site/Style Library/send-docs-customaction.js" runat="server"/>-->
<!--SPM:<SharePoint:ScriptLink language="javascript" name="~site/Style Library/send-docs-ribbon.js" runat="server"/>-->
<!--SPM:<SharePoint:ScriptLink language="javascript" name="~site/Style Library/SendDocument.js" runat="server"/>-->
<!--SPM:<SharePoint:ScriptLink language="javascript" name="~site/Style Library/send-docs-main.js" runat="server"/>-->
</div>
