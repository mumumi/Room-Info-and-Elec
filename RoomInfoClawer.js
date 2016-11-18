//??????select?onchange??doPostBack???select??????????js?????

function sleep(numberMillis) {
    var now = new Date();
    var exitTime = now.getTime() + numberMillis;
    while (true) {
        now = new Date();
        if (now.getTime() > exitTime)
            return;
    }
}

function doSelect(selectId,selectIndex) {
    document.getElementById(selectId).options[selectIndex].selected = true;
    document.getElementById(selectId).onchange();
	document.getElementById(selectId).focus();
	document.getElementById('BianHao').focus();
    return document.getElementById(selectId)[document.getElementById(selectId).selectedIndex].text;
}

var ta = document.createElement("textarea");
ta.setAttribute("type", "text");
ta.setAttribute("id", "test");
ta.setAttribute("style", "height:500px");
var div = document.getElementById("form1");
div.appendChild(ta);

for (var i = 0; i < document.getElementById('DDLSheQu').options.length; i++) {
    document.getElementById('DDLSheQu').options[i].selected = true;
    setTimeout('__doPostBack(\'DDLSheQu\',\'\')', 0);
    //sleep(1000);
    var DDLSheQuName = document.getElementById('DDLSheQu')[document.getElementById('DDLSheQu').selectedIndex].innerText;

    for (var j = 0; j < document.getElementById('DDLLouYu').options.length; j++) {
        document.getElementById('DDLLouYu').options[j].selected = true;
        setTimeout('__doPostBack(\'DDLLouYu\',\'\')', 0);
        //sleep(1000);
        var DDLLouYuName = DDLSheQuName + document.getElementById('DDLLouYu')[document.getElementById('DDLLouYu').selectedIndex].innerText;

        for (var k = 0; k < document.getElementById('DDLDanYuan').options.length; k++) {
            document.getElementById('DDLDanYuan').options[k].selected = true;
            setTimeout('__doPostBack(\'DDLDanYuan\',\'\')', 0);
            //sleep(1000);
            var DDLDanYuanName = DDLLouYuName + document.getElementById('DDLDanYuan')[document.getElementById('DDLDanYuan').selectedIndex].innerText + "???";

            for (var l = 0; l < document.getElementById('DDLLouCeng').options.length; l++) {
                document.getElementById('DDLLouCeng').options[l].selected = true;
                setTimeout('__doPostBack(\'DDLLouCeng\',\'\')', 0);
                //sleep(1000);
                var DDLLouCengName = DDLDanYuanName + document.getElementById('DDLLouCeng')[document.getElementById('DDLLouCeng').selectedIndex].innerText + "??";

                for (var m = 0; m < document.getElementById('DDLFangJian').options.length; m++) {
                    document.getElementById('DDLFangJian').options[m].selected = true;
                    setTimeout('__doPostBack(\'DDLFangJian\',\'\')', 0);
                    //sleep(1000);
                    var DDLFangJian = DDLLouCengName + document.getElementById('DDLFangJian')[document.getElementById('DDLFangJian').selectedIndex].innerText + "??";

                    document.getElementById("test").value += document.getElementById('BianHao').innerText.trim() + "," + DDLFangJian + ";";

                }
            }
        }
    }
}

//+++++++++++++++++++++++++++++++++++++++++++++




var DDLFangJian = document.getElementById('DDLFangJian');
for (var m = 0; m < DDLFangJian.options.length; m++) {
    DDLFangJian.options[m].selected = true;
    setTimeout('__doPostBack(\'DDLFangJian\',\'\')', 0);
    sleep(1000);
    var DDLFangJian = DDLFangJian.options[DDLFangJian.selectedIndex].innerText + "??";
//alert(DDLFangJian)
    document.getElementById("test").value += document.getElementById('BianHao').innerText.trim() + "," + DDLFangJian + ";";

}


for (var m = 0; m < document.getElementById('DDLFangJian').options.length; m++) {
    document.getElementById('DDLFangJian').options[m].selected = true;
    setTimeout('__doPostBack(\'DDLFangJian\',\'\')', 0);
    //sleep(1000);
    var DDLFangJian = document.getElementById('DDLFangJian').options[document.getElementById('DDLFangJian').selectedIndex].innerText + "??";
//alert(DDLFangJian)
    document.getElementById("test").value += document.getElementById('BianHao').innerText.trim() + "," + DDLFangJian + ";";

}

            for (var l = 0; l < document.getElementById('DDLLouCeng').options.length; l++) {
                document.getElementById('DDLLouCeng').options[l].selected = true;
                __doPostBack('DDLLouCeng','');
				//document.getElementById('DDLFangJian').focus();
                //sleep(1000);
                var DDLLouCengName = document.getElementById('DDLLouCeng')[document.getElementById('DDLLouCeng').selectedIndex].innerText + "??";

                for (var m = 0; m < document.getElementById('DDLFangJian').options.length; m++) {
                    document.getElementById('DDLFangJian').options[m].selected = true;
                    __doPostBack('DDLFangJian','');
				//document.getElementById('BianHao').focus();
                    //sleep(1000);
                    var DDLFangJian = DDLLouCengName + document.getElementById('DDLFangJian')[document.getElementById('DDLFangJian').selectedIndex].innerText + "??";

                    document.getElementById("test").value += document.getElementById('BianHao').innerText.trim() + "," + DDLFangJian + ";";

                }
            }

            for (var l = 0; l < document.getElementById('DDLLouCeng').options.length; l++) {
                document.getElementById('DDLLouCeng').options[l].selected = true;
                document.write("<script type=\"text/javascript\" >__doPostBack(\'DDLLouCeng\',\'\')</script>"); 
				
				//document.getElementById('DDLFangJian').focus();
                //sleep(10000);
                var DDLLouCengName = document.getElementById('DDLLouCeng')[document.getElementById('DDLLouCeng').selectedIndex].innerText + "??";

                    document.getElementById('DDLFangJian').options[2].selected = true;
                    document.write("<script type=\"text/javascript\" >__doPostBack(\'DDLFangJian\',\'\')</script>"); 
				//document.getElementById('BianHao').focus();
                    //sleep(100);
                    var DDLFangJian = DDLLouCengName + document.getElementById('DDLFangJian')[document.getElementById('DDLFangJian').selectedIndex].innerText + "??";

                    document.getElementById("test").value += document.getElementById('BianHao').innerText.trim() + "," + DDLFangJian + ";";

                
            }
			
			document.all.DDLLouCeng.options[1].selected = true;
			//document.all.DDLLouCeng.onchange();
			document.write("<script type=\"text/javascript\" >__doPostBack(\'DDLLouCeng\',\'\')</script>"); 
			
			//__doPostBack('DDLLouCeng','');
			//alert();
			sleep(1000);
			document.all.DDLFangJian.options[2].selected = true;
			alert(document.all.DDLFangJian[document.all.DDLFangJian.selectedIndex].innerText + "??");
			
			document.all.DDLLouCeng.options[1].selected = true;
			//document.all.DDLLouCeng.onchange();
			var script=document.createElement("script");  
            script.async = true;
			script.setAttribute("type", "text/javascript");  
			script.setAttribute("id", "mys");  
			script.text =  "__doPostBack(\'DDLLouCeng\',\'\')";
			var div = document.getElementById("UpdatePanel1");
			document.all.UpdatePanel1.appendChild(script);
			//__doPostBack('DDLLouCeng','');
			//alert();
			sleep(1000);
			document.all.DDLFangJian.options[2].selected = true;
			alert(document.all.DDLFangJian[document.all.DDLFangJian.selectedIndex].innerText + "??");
			
			
			
			document.all.DDLLouCeng.options[1].selected = true;
			document.getElementById("mys").text =  "__doPostBack(\'DDLLouCeng\',\'\')";
			
			
			document.all.DDLLouCeng.options[1].selected = true;
			//document.all.DDLLouCeng.onchange();
			var script=document.createElement("script");  
			script.async = true;
			script.setAttribute("type", "text/javascript");  
			script.setAttribute("id", "mys");  
			script.text =  "function doSelect(selectId,selectIndex) {document.getElementById(selectId).options[selectIndex].selected = true; document.getElementById(selectId).onchange(); return document.getElementById(selectId)[document.getElementById(selectId).selectedIndex].text;}";
			var div = document.getElementById("UpdatePanel1");
			document.all.UpdatePanel1.appendChild(script);
			//__doPostBack('DDLLouCeng','');
			//alert();
			sleep(1000);
			document.all.DDLFangJian.options[2].selected = true;
			alert(document.all.DDLFangJian[document.all.DDLFangJian.selectedIndex].innerText + "??");
			
			function doSelect(selectId,selectIndex) {document.getElementById(selectId).options[selectIndex].selected = true; __doPostBack(selectId,\'\'); return document.getElementById(selectId)[document.getElementById(selectId).selectedIndex].text;}
			

			var script=document.createElement("script"); 
			script.setAttribute("type", "text/javascript");  			
			script.setAttribute("src", "http://code.jquery.com/jquery-1.4.1.min.js");
			document.all.UpdatePanel1.appendChild(script);
			
			
			
<script type="text/javascript">
window.onload = function() {
var path = "C:\\test.txt";
            var fso = new ActiveXObject("Scripting.FileSystemObject");
            var s = fso.CreateTextFile(path, true);
            var content = document.getElementById('BianHao').innerText.trim();
            s.WriteLine(content);
            s.Close();
};
</script>


document.getElementById('DDLSheQu').options[3].selected =true
setTimeout('__doPostBack(\'DDLSheQu\',\'\')', 0)
document.getElementById('BianHao').innerText.trim()

document.getElementById("test").value+="abc;";


document.getElementById("divList").innerHTML+='<li><input type="text" name="test" size="40" /></li><br />';
