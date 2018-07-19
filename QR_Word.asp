
<%
	'// Force declaration of all variables
	Option Explicit
	'// Allow response.redirect
	Response.Buffer = True

	'// Variable declaration
	Dim objConn							'// Stores the ADODB connection object
	Dim strSQL							'// Stores the SQL script to be executed
    Dim QRCode
    Dim rsMerchant


QRCode                   = Request.querystring("reference")
rsMerchant               = Request.QueryString("name")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!-- //////////////////////////////////////////////////////////////////////// -->
<!-- // ATM/SST/POS Operations Information Management Systems              // -->
<!-- //////////////////////////////////////////////////////////////////////// -->
<!-- // Filename    : QR_Word.asp                                       // -->
<!-- // Purpose     : generate word ducument                                           // -->
<!-- // Version     : 1.00.00                                              // -->
<!-- // Revisions   : None                                                 // -->
<!-- // Date        : June , 2018                                        // -->
<!-- // Author      : Michael Gasa                                           // -->
<!-- //                         // -->
<!-- //////////////////////////////////////////////////////////////////////// -->
<!--#include file="../Menu/ServerVariable.asp"-->
<html xmlns ="http://www.w3.org/1999/xhtml" >
<head>
<%
Response.ContentType = "application/msword"
Response.AddHeader "Content-Disposition", "attachment;filename=wordfile.doc"
Response.write("<html " & _
"xmlns:o='urn:schemas-microsoft-com:office:office' " & _
"xmlns:w='urn:schemas-microsoft-com:office:word'" & _
"xmlns='http://www.w3.org/TR/REC-html40'>")
%>
<title>QR Code </title>

        <title>Testing QR code</title>
        <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.11.2/jquery.min.js"></script>
        <script type="text/javascript">
            function generateBarCode() {
                var nric = '" & QRCode & "'
                var url = 'https://api.qrserver.com/v1/create-qr-code/?data=' + nric + '&amp;size=162x162';
                $('#barcode').attr('src', url);
            }
        </script>
        <style type="text/css">
            #barcode {
                width: 328px;
                height: 328px;
            }

img.center {
    display: block;
    margin: 0 auto;
}
        </style>


</head>
<body style="margin-top:0; margin-centre:0; margin-right:0">
    <h1>
<center>   <%response.Write (rsMerchant) %></center> </h1> <br />
<center>
  
      <img id='barcode' 
            src="https://api.qrserver.com/v1/create-qr-code/?data=<%=QRCode%>&amp;size=162x162" 
            alt="" 
            title="Merchant" align="middle" />&nbsp;

    <br />
    
    <h1>

         
           <br />
         <%response.Write (QRCode) %>

     
</h1>
</center>



</body>
</html>
