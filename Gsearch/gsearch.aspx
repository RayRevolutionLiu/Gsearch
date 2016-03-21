<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="gsearch.aspx.cs" Inherits="gsearch" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
            <!-- tabmenu -->
        <div class="tabmenuT2block">
            <div class="fixwidth twocol">
                <div class="textcenter">
                    <!-- 選單位置:left:置左、right:置右、center:置中 -->
                    <span class="tabmenuT2Btn"><a href="index.aspx" target="_self">年度報表</a></span>
                    <span class="tabmenuT2Btn"><a href="list.aspx" target="_self">最常使用者</a></span>
                    <span class="tabmenuT2Btn  tabmenuT2BtnCurrent"><a href="gsearch.aspx" target="_self">查詢條件</a></span>
                    <span class="tabmenuT2Btn"><a href="importexcel.aspx" target="_self">上傳檔案</a></span>
                </div>
                <!-- left -->
            </div>
            <!-- twocol -->
        </div>
        <!-- tabmenuT2block -->
        <!-- tabmenu -->
                <div class="fixwidth">
	<div class="gentable font-size3 margin10T">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <th width="18%"align="right"><div class="titlebackicon">資料庫名稱</div></th>
    <td width="82%" class="width85" colspan="3">
        <select multiple="multiple" id="main_dbtype" >
<%--      <option value="1">ACM</option>--%>
    </select></td>
    </tr>
  <tr>
    <th align="right"><div class="titlebackicon">單位</div></th>
    <td colspan="3"><select multiple="multiple" id="orgcd">
<%--      <option value="1">01院部</option>--%>
    </select></td>
    </tr>
  <tr>
    <th width="18%" align="right"><div class="titlebackicon">刊名</div></th>
    <td width="32%"><input type="text" class="inputex width90" id="main_title"  maxlength="150"/></td>
    <th width="18%" align="right"><div class="titlebackicon">篇名</div></th>
    <td width="32%"><input type="text" class="inputex width90" id="main_chapter" maxlength="400" /></td>
  </tr>
  <tr>
    <th align="right"><div class="titlebackicon">工號</div></th>
    <td><input type="text" class="inputex width90" id="main_empno" maxlength="6"/></td>
    <th align="right"><div class="titlebackicon">部門代碼</div></th>
    <td><input type="text" class="inputex width90" id="main_deptcd" maxlength="10" /></td>
  </tr>
  <tr>
    <th align="right"><div class="titlebackicon">姓名</div></th>
    <td><input type="text" class="inputex width90" id="main_cname" maxlength="10" /></td>
    <th align="right"></th>
    <td>&nbsp;</td>
  </tr>
    <tr>
    <th align="right"><div class="titlebackicon">期間</div></th>
    <td colspan="2"><input type="text" class="inputex datepicker" id="main_startdate"/> ~ <input type="text" class="inputex datepicker" id="main_enddate"/></td>
    <td align="right"><span style="margin-right:22px;"><input type="button" value="查詢" class="genbtn searchbtn" id="searchbtn" /><input type="button" value="清除" class="genbtn searchbtn" id="clearbtn" /></span></td>
    </tr>
</table>
</div>

</div>

<div id="searchoutcome">
<div class="twocol margin15TB padding5RL">
<div class="left"><h3>搜尋結果<span id="totalCountNum"></span></h3></div>
<div class="right">
    <input type="button" value="匯出CSV" class="genbtn searchbtn" id="btnExportExcel" />
</div>
</div>
<div class="stripeMe" style="padding:0 5px;">
<table width="100%" id="createTable">
 <%-- <tr>
    <th>資料庫名稱</th>
    <th>日期</th>
    <th>單位</th>
    <th>部門名稱</th>
    <th>部門代碼</th>
    <th>姓名</th>
    <th>工號</th>
    <th>瀏覽次數</th>
    <th>下載次數</th>
    <th>登入次數</th>
    <th>檢索次數</th>
    <th>信箱</th>
    <th>刊名</th>
    <th>篇名</th>
  </tr>
  <tr>
    <td>Digitimes</td>
    <td>2015/09/28</td>
    <td>17資科</td>
    <td>網際網路應用系統整合二部</td>
    <td>SF200</td>
    <td>詹將佑</td>
    <td>529130</td>
    <td>12</td>
    <td>&nbsp;</td>
    <td>33&nbsp;</td>
    <td>&nbsp;</td>
    <td>529130@itri.org.tw</td>
    <td>Weimar publics/Weimar 
        <br />
        subjects</td>
    <td>rethinking the political culture of 
        <br />
        Germany in the 1920s</td>
  </tr>--%>
</table>
</div>
</div><!-- searchoutcome -->

<%--dialog--%>
<div class="dialog ui-dialog-titlebar lineheight03">
    <div id="processbar" style="display:none;">
        <img src="images/loading.gif"/>
    </div>
    <div id="sendmsg"></div>
</div>

    <script type="text/javascript">
        var intervalID;

        $(document).ready(function () {
            $(window).scroll(function () {
                if (($(document).height() - $(window).height()) / $(window).scrollTop() < 3) {
                    dataArr.flag++;
                    asyncGenTable('','');
                }
                
                //if ($(window).scrollTop() == $(document).height() - $(window).height()) {
                //    if ($(".pagenum:last").val() <= $(".rowcount").val()) {
                //        var pagenum = parseInt($(".pagenum:last").val()) + 1;
                //        getresult('getresult.php?page=' + pagenum);
                //    }
                //}
            });

            $.ajax({
                type: "POST",
                url: "handler/SearchConditions.ashx",
                async:false,
                //data: {
                //    type: "dbtype"
                //},
                error: function (xhr, textStatus) {
                    console.log(xhr.responseText);
                    alert("查詢條件連繫錯誤");
                },
                //dataType: "text",
                success: function (response) {
                    var dbtypeSTR = "";
                    var orgcdSTR = "";
                    $("#processbar").show();
                    for (var i = 0; i < response.length; i++) {
                        if (response[i].type_group == "main_dbtype") {
                            //資料庫類別
                            dbtypeSTR += "<option value='" + response[i].value + "'>" + response[i].cname + "</option>";
                        }
                        if (response[i].type_group == "orgcd") {
                            orgcdSTR += "<option value='" + response[i].value + "'>" + response[i].cname + "</option>";
                        }
                    }

                    $("#main_dbtype").append(dbtypeSTR);
                    $("#orgcd").append(orgcdSTR);
                    $("#processbar").hide();
                }
            });


            $("#searchbtn").click(function () {
                if (Searchcheck() != "") {
                    alert(Searchcheck());
                    return false;
                }
                asyncGenTable('btn','');
            });

            $("#clearbtn").click(function () {
                location.href = 'gsearch.aspx';
            });

            //下載使用者上傳的檔案
            $(document).on("click", "#btnExportExcel", function (event) {
                $(".dialog").dialog("open").dialog({ width: 240, height: 70 })
                $("#sendmsg").text("下載中請稍後...");
                $("#processbar").show();

                var selected = dataArr.selected;
                var selectedOrg = dataArr.selectedOrg;
                var main_title = $("#main_title").val();
                var main_chapter = $("#main_chapter").val();
                var main_empno = $("#main_empno").val();
                var main_cname = $("#main_cname").val();
                var main_deptcd =  $("#main_deptcd").val();
                var main_startdate =  $("#main_startdate").val();
                var main_enddate =  $("#main_enddate").val();
                //var pageSize = dataArr.pageSize;
                //var flag =  dataArr.flag;
                var sortcolumn =  dataArr.sortcolumn;
                var orderBy = dataArr.orderBy;
                var Export = "0";

                var form = document.createElement('form');
                var iframe = document.createElement('iframe');
                var selectedIn = document.createElement('input');
                var selectedOrgIn = document.createElement('input');
                var main_titleIn = document.createElement('input');
                var main_chapterIn = document.createElement('input');
                var main_empnoIn = document.createElement('input');
                var main_cnameIn = document.createElement('input');
                var main_deptcdIn = document.createElement('input');
                var main_startdateIn = document.createElement('input');
                var main_enddateIn = document.createElement('input');
                var sortcolumnIn = document.createElement('input');
                var orderByIn = document.createElement('input');
                var ExportIn = document.createElement('input');
                var UUID = generateUUID();

                //以下是找不到辦法,只能一個一個加
                selectedIn.setAttribute("id", "selectedIn");
                selectedIn.setAttribute("name", "selectedIn");
                selectedIn.setAttribute("value", selected);

                selectedOrgIn.setAttribute("id", "selectedOrgIn");
                selectedOrgIn.setAttribute("name", "selectedOrgIn");
                selectedOrgIn.setAttribute("value", selectedOrg);

                main_titleIn.setAttribute("id", "main_titleIn");
                main_titleIn.setAttribute("name", "main_titleIn");
                main_titleIn.setAttribute("value", main_title);

                main_chapterIn.setAttribute("id", "main_chapterIn");
                main_chapterIn.setAttribute("name", "main_chapterIn");
                main_chapterIn.setAttribute("value", main_chapter);

                main_empnoIn.setAttribute("id", "main_empnoIn");
                main_empnoIn.setAttribute("name", "main_empnoIn");
                main_empnoIn.setAttribute("value", main_empno);

                main_cnameIn.setAttribute("id", "main_cnameIn");
                main_cnameIn.setAttribute("name", "main_cnameIn");
                main_cnameIn.setAttribute("value", main_cname);

                main_deptcdIn.setAttribute("id", "main_deptcdIn");
                main_deptcdIn.setAttribute("name", "main_deptcdIn");
                main_deptcdIn.setAttribute("value", main_deptcd);

                main_startdateIn.setAttribute("id", "main_startdateIn");
                main_startdateIn.setAttribute("name", "main_startdateIn");
                main_startdateIn.setAttribute("value", main_startdate);

                main_enddateIn.setAttribute("id", "main_enddateIn");
                main_enddateIn.setAttribute("name", "main_enddateIn");
                main_enddateIn.setAttribute("value", main_enddate);

                sortcolumnIn.setAttribute("id", "sortcolumnIn");
                sortcolumnIn.setAttribute("name", "sortcolumnIn");
                sortcolumnIn.setAttribute("value", sortcolumn);

                orderByIn.setAttribute("id", "orderByIn");
                orderByIn.setAttribute("name", "orderByIn");
                orderByIn.setAttribute("value", orderBy);

                ExportIn.setAttribute("id", "ExportIn");
                ExportIn.setAttribute("name", "ExportIn");
                ExportIn.setAttribute("value", Export);

                iframe.setAttribute("name", "postiframe");
                iframe.setAttribute("id", "postiframe");
                iframe.setAttribute("style", "display: none");

                form.appendChild(iframe);
                form.appendChild(selectedIn);
                form.appendChild(selectedOrgIn);
                form.appendChild(main_titleIn);
                form.appendChild(main_chapterIn);
                form.appendChild(main_empnoIn);
                form.appendChild(main_cnameIn);
                form.appendChild(main_deptcdIn);
                form.appendChild(main_startdateIn);
                form.appendChild(main_enddateIn);
                form.appendChild(sortcolumnIn);
                form.appendChild(orderByIn);
                form.appendChild(ExportIn);
                $("body").append(form);

                form.setAttribute("action", "handler/filedownload.ashx");
                form.setAttribute("method", "post");
                form.setAttribute("enctype", "multipart/form-data");
                form.setAttribute("encoding", "multipart/form-data");
                form.setAttribute("target", "postiframe");
                form.setAttribute("style", "display: none");
                form.setAttribute("name", UUID);
                form.submit();

                intervalID = window.setInterval(KeepAskSessionStatus, 2000);
            });
        });

        var dataArr = {
            selected: '',
            selectedOrg: '',
            pageSize: 100,//每個區塊要幾筆顯示
            frameHeight: $(window).height(),
            //totalframe: $(document).height() % $(window).height() != 0 ? $(document).height() / $(window).height() + 1 : $(document).height() / $(window).height(),
            flag: 1,
            orderBy: 'ASC',
            sortcolumn: 'main_id'
        }

        function asyncGenTable(ev,sortcolumn) {
            if (ev != '') {
                $("#createTable").children().remove();
                dataArr.flag = 1;
                if (sortcolumn != "") {
                    dataArr.sortcolumn = sortcolumn;
                }

                if (dataArr.orderBy == 'ASC' && sortcolumn != "") {
                    dataArr.orderBy = 'DESC';
                }
                else if (dataArr.orderBy == 'DESC' && sortcolumn != "") {
                    dataArr.orderBy = 'ASC';
                }
            }
            
            $.ajax({
                type: "POST",
                url: "handler/GenTable.ashx",
                async: false,
                data: {
                    selected: dataArr.selected,
                    selectedOrg: dataArr.selectedOrg,
                    main_title: $("#main_title").val(),
                    main_chapter: $("#main_chapter").val(),
                    main_empno: $("#main_empno").val(),
                    main_cname: $("#main_cname").val(),
                    main_deptcd: $("#main_deptcd").val(),
                    main_startdate: $("#main_startdate").val(),
                    main_enddate: $("#main_enddate").val(),
                    pageSize: dataArr.pageSize,
                    //totalHeight: dataArr.totalHeight,
                    //frameHeight: dataArr.frameHeight,
                    flag: dataArr.flag,
                    sortcolumn: dataArr.sortcolumn,
                    orderBy: dataArr.orderBy,
                    Export: "1"
                },
                error: function (xhr, textStatus) {
                    alert("表格產生錯誤" + xhr.responseText);
                },
                //dataType: "text",
                success: function (response) {
                    if (response.length > 0) {
                        var headInnerHTML = '<tr>';
                        headInnerHTML += '<th><a href="javascript:void(0);" onclick="asyncGenTable(\'btn\',\'main_dbtype\');">資料庫名稱</a></th>';
                        headInnerHTML += '<th><a href="javascript:void(0);" onclick="asyncGenTable(\'btn\',\'main_datetime\');">日期</a></th>';
                        headInnerHTML += '<th><a href="javascript:void(0);" onclick="asyncGenTable(\'btn\',\'main_orgcname\');">單位</a></th>';
                        headInnerHTML += '<th><a href="javascript:void(0);" onclick="asyncGenTable(\'btn\',\'main_deptcdcname\');">部門名稱</a></th>';
                        headInnerHTML += '<th><a href="javascript:void(0);" onclick="asyncGenTable(\'btn\',\'main_deptcd\');">部門代碼</a></th>';
                        headInnerHTML += '<th><a href="javascript:void(0);" onclick="asyncGenTable(\'btn\',\'main_cname\');">姓名</a></th>';
                        headInnerHTML += '<th><a href="javascript:void(0);" onclick="asyncGenTable(\'btn\',\'main_empno\');">工號</a></th>';
                        headInnerHTML += '<th><a href="javascript:void(0);" onclick="asyncGenTable(\'btn\',\'main_viewC\');">瀏覽次數</a></th>';
                        headInnerHTML += '<th><a href="javascript:void(0);" onclick="asyncGenTable(\'btn\',\'main_downloadC\');">下載次數</a></th>';
                        headInnerHTML += '<th><a href="javascript:void(0);" onclick="asyncGenTable(\'btn\',\'main_logC\');">登入次數</a></th>';
                        headInnerHTML += '<th><a href="javascript:void(0);" onclick="asyncGenTable(\'btn\',\'main_indexC\');">檢索次數</a></th>';
                        headInnerHTML += '<th>信箱</th><th>刊名</th><th>篇名</th></tr>';

                        var bodyInnerHTML = ''
                        for (var i = 0; i < response.length; i++) {
                            bodyInnerHTML += '<tr>';
                            bodyInnerHTML += '<td>' + response[i].main_dbtype + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_datetime + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_orgcname + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_deptcdcname + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_deptcd + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_cname + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_empno + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_viewC + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_downloadC + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_logC + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_indexC + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_email + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_title + '</td>';
                            bodyInnerHTML += '<td>' + response[i].main_chapter + '</td>';
                            bodyInnerHTML += '</tr>';
                        }
                        if (dataArr.flag == 1) {
                            //只有第一次需要把表頭binding
                            var act1 = TweenLite.to($("#searchoutcome"), 0.8, { display: "block", opacity: 1, y: 10 });
                            $("#createTable").append(headInnerHTML);
                        }

                        $("#createTable").append(bodyInnerHTML);
                        $("#totalCountNum").text("(" + response[0].totalcount + "筆資料)");
                        $("#btnExportExcel").show();
                    }
                    else if (ev != '') {
                        var act1 = TweenLite.to($("#searchoutcome"), 0.8, { display: "block", opacity: 1, y: 10 });
                        $("#createTable").append(headInnerHTML);
                        $("#totalCountNum").text("(0筆資料)");
                        $("#createTable").append('<tr><td>查詢無資料</td></tr>');
                        $("#btnExportExcel").hide();
                    }
                }
            });
        }

        function generateUUID() {
            var d = new Date().getTime();
            var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
                var r = (d + Math.random() * 16) % 16 | 0;
                d = Math.floor(d / 16);
                return (c == 'x' ? r : (r & 0x3 | 0x8)).toString(16);
            });
            return uuid;
        }

        function Searchcheck() {
            var check = '請輸入查詢條件';

            //資料庫類別
            var len = $("#main_dbtype").get(0).options.length;
            dataArr.selected = '';
            for (var i = 0; i < len; i++) {
                if ($("#main_dbtype").get(0).options[i].selected) {
                    dataArr.selected += $("#main_dbtype").get(0).options[i].value + ",";
                    check = '';
                }
            }
            dataArr.selected = dataArr.selected.substr(0, dataArr.selected.length - 1);


            //單位
            var lenOrg = $("#orgcd").get(0).options.length;
            dataArr.selectedOrg = '';
            for (var i = 0; i < lenOrg; i++) {
                if ($("#orgcd").get(0).options[i].selected) {
                    dataArr.selectedOrg += $("#orgcd").get(0).options[i].value + ",";
                    check = '';
                }
            }
            dataArr.selectedOrg = dataArr.selectedOrg.substr(0, dataArr.selectedOrg.length - 1);

            //刊名
            if ($("#main_title").val() != "") check = '';
            //篇名
            if ($("#main_chapter").val() != "") check = '';
            //工號
            if ($("#main_empno").val() != "") check = '';
            //姓名
            if ($("#main_cname").val() != "") check = '';
            //部門代碼
            if ($("#main_deptcd").val() != "") check = '';
            //期間
            if ($("#main_startdate").val() != "" && $("#main_enddate").val() != "") {
                check = '';
            }
            else if ($("#main_startdate").val() != "" && $("#main_enddate").val() == "") {
                check = '請輸入結束日期';
            }
            else if ($("#main_startdate").val() == "" && $("#main_enddate").val() != "") {
                check = '請輸入開始日期';
            }

            return check;
        }


        function KeepAskSessionStatus() {
            $.ajax({
                type: "POST",
                url: "handler/KeepAskSessionStatus.ashx",
                error: function (xhr, textStatus) {
                    console.log(xhr.responseText);
                    alert("下載狀態連繫錯誤");
                    window.clearInterval(intervalID);
                },
                //dataType: "text",
                success: function (response) {
                    if (response == "success") {
                        window.clearInterval(intervalID);
                        $(".dialog").dialog("close");
                    }
                }
            });
        }


        //function ajaxAlertFun(msg) {
        //    //form submit出去後,即可清除
        //    $("form[name='" + UUID + "']").remove();
        //    $(".ui-dialog-titlebar").show();
        //    $("#processbar").hide();

        //    if (msg.indexOf('error') > -1) {
        //        //error_錯誤訊息
        //        //$(".dialog").dialog("open");
        //        $(".dialog").dialog('option', 'title', "警告");
        //        $("#sendmsg").text(msg.substring(msg.indexOf('_') + 1, msg.length));
        //    }
        //    else {
        //        $(".dialog").dialog("close");
        //    }

        //}
    </script>
</asp:Content>

