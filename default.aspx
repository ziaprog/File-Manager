
<%@ Page Language="vb" AutoEventWireup="false" Inherits="WebFileManager._default" CodeFile="default.aspx.vb" %>
<HTML>
	<HEAD>
		<TITLE>WebFileManager Browsing
			<%=CurrentWebPath%>
		</TITLE>
		<META content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<SCRIPT language="JavaScript">	
		// create new folder
		function newfolder() {
			if (document.forms[0].elements['<%=TargetFolderTag%>'].value == '') {
				alert('No folder name was provided. Enter a new folder name in the textbox.');
				document.forms[0].elements['<%=TargetFolderTag%>'].focus();
				return;
			}
			document.forms[0].action.value = 'newfolder';
			document.forms[0].submit();
		}
		
		function upload() {
			if (document.forms[0].elements['fileupload'].value == '') {
				alert('No upload file was provided. Select a file to upload.');
				document.forms[0].elements['fileupload'].focus();
				return;
			}
			document.forms[0].action.value = 'upload';
			document.forms[0].submit();
		}
		
		// check or uncheck all file checkboxes
		function checkall(ctl) {
			for (var i = 0; i < document.forms[0].elements.length; i++) {
			    if (document.forms[0].elements[i].name.indexOf('<%=CheckboxTag%>') > -1) { 
			        document.forms[0].elements[i].checked = ctl.checked;
			        }
				}
		}
		
		// confirm file list and target folder
		function confirmfiles(sAction) {
			var nMarked = 0;
			var sTemp = '';
			for (var i = 0; i < document.forms[0].elements.length; i++) {
				if (document.forms[0].elements[i].checked && 
				 document.forms[0].elements[i].name.indexOf('<%=CheckboxTag%>') > -1) { 
					if (sAction == 'rename') {
						var sFilename = '';
            			var sNewFilename = '';
						sFilename = document.forms[0].elements[i].name;
						sFilename = sFilename.replace('<%=CheckboxTag%>','');
						sNewFilename = prompt('Enter new name for ' + sFilename, sFilename);
						if (sNewFilename != null) {
						    document.forms[0].elements[i].name = document.forms[0].elements[i].name + '<%=RenameTag%>' + sNewFilename;
						}
					}
					nMarked = nMarked + 1;
				}
			}
			if (nMarked == 0) {
				alert('No items selected. To select items, use the checkboxes on the left.');
				return;
			}
			sTemp = 'Are you sure that you want to ' + sAction + ' the ' + nMarked + ' checked item(s)?'
			if (sAction == 'copy' || sAction == 'move') {
				sTemp = 'Are you sure you want to ' + sAction + ' the ' + nMarked + ' checked item(s) to the "' + document.forms[0].elements['<%=TargetFolderTag%>'].value + '" folder?'
				if (document.forms[0].elements['<%=TargetFolderTag%>'].value == '') {
				    document.forms[0].elements['<%=TargetFolderTag%>'].focus();
					alert('No destination folder provided. Enter a folder name.');
					return;
				}
			}
			var confirmed = false;
			if (sAction == 'copy' || sAction == 'rename') {
			    confirmed = true;
			} else {
			    confirmed = confirm(sTemp);
			}

			if (confirmed) { 
				document.forms[0].action.value = sAction;
				document.forms[0].submit();
			}

		}
		</SCRIPT>
		<STYLE type="text/css">
            IMG { border: none; }
            BODY { FONT-FAMILY: Verdana, Arial, Helvetica; FONT-SIZE: 70%; margin: 4px 4px 4px 4px;}
            TD { FONT-FAMILY: Verdana, Arial, Helvetica; FONT-SIZE: 70%; }
            TH { FONT-FAMILY: Verdana, Arial, Helvetica; FONT-SIZE: 70%; background-color: #eeeeee;}
            .Header { background-color: lightyellow; }
            .Error { color:red; font-weight: bold; }
		</STYLE>
	</HEAD>
	<BODY>
		<%HandleAction()%>
		<FORM action="<%=ScriptName%>" method=post enctype=multipart/form-data>
			<INPUT type=hidden name="<%=ActionTag%>">
			<%WriteTable()%>
		</FORM>
	</BODY>
</HTML>
