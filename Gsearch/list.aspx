<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="list.aspx.cs" Inherits="list" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
                <!-- tabmenu -->
        <div class="tabmenuT2block">
            <div class="fixwidth twocol">
                <div class="textcenter">
                    <!-- 選單位置:left:置左、right:置右、center:置中 -->
                    <span class="tabmenuT2Btn"><a href="index.aspx" target="_self">年度報表</a></span>
                    <span class="tabmenuT2Btn tabmenuT2BtnCurrent"><a href="list.aspx" target="_self">最常使用者</a></span>
                    <span class="tabmenuT2Btn"><a href="gsearch.aspx" target="_self">查詢條件</a></span>
                    <span class="tabmenuT2Btn"><a href="importexcel.aspx" target="_self">上傳檔案</a></span>
                </div>
                <!-- left -->
            </div>
            <!-- twocol -->
        </div>
        <!-- tabmenuT2block -->
    <div style="text-align: center; padding-top:15px;" align="center">
        選擇資料庫：<select id="selectDB"></select>
        選擇類型：<select id="selectCountType">
            <option value="">全部</option>
            <option value="main_downloadC">下載數</option>
            <option value="main_viewC">瀏覽數</option>
            <option value="main_indexC">檢索數</option>
            <option value="main_logC">登入數</option>
            </select>
        選擇年度：<select id="selectYY"></select>
        <input type="button" value="匯出" class="genbtn searchbtn" id="ExportAllbtn" />
    </div>
    <div id="genDIV" class="stripeMe margin15TB" style="padding:0 10px;">
    </div>

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
                $.ajax({
                    type: "POST",
                    url: "handler/SearchConditions.ashx",
                    async: false,
                    error: function (xhr, textStatus) {
                        console.log(xhr.responseText);
                        alert("資料庫下拉選單連繫錯誤");
                    },
                    //dataType: "text",
                    success: function (response) {
                        var optonSTR = '<option value="">全部</option>';
                        if (response.length > 0) {  
                            for (var i = 0; i < response.length; i++) {
                                if (response[i].type_group == "main_dbtype") {
                                    //資料庫類別
                                    optonSTR += "<option value='" + response[i].value + "'>" + response[i].cname + "</option>";
                                }
                            }
                        }
                        $("#selectDB").append(optonSTR);
                    }
                });

                $.ajax({
                    type: "POST",
                    url: "handler/getYYgroupBy.ashx",
                    async: false,
                    error: function (xhr, textStatus) {
                        console.log(xhr.responseText);
                        alert("年分下拉選單連繫錯誤");
                    },
                    //dataType: "text",
                    success: function (response) {
                        $("#selectYY").children().remove();
                        var optonSTR = '<option value="">全部</option>';
                        if (response.length > 0) {
                            for (var i = 0; i < response.length; i++) {
                                optonSTR += "<option value='" + response[i].value + "'>" + response[i].cname + "</option>";
                            }
                        }
                        $("#selectYY").append(optonSTR);
                    }
                });

                bindTable();


                $("#selectDB").change(function () {
                    bindTable();
                });

                $("#selectYY").change(function () {
                    bindTable();
                });

                $("#selectCountType").change(function () {
                    bindTable();
                    if ($("#selectCountType").val() != '') {
                        $("#thCount").text($("#selectCountType :selected").text());
                    }
                });

                $(document).on("click", "#ExportAllbtn", function (event) {
                    $(".dialog").dialog("open").dialog({ width: 240, height: 70 })
                    $("#sendmsg").text("下載中請稍後...");
                    $("#processbar").show();

                    var selectDB = $("#selectDB").val();
                    var selectCountType = $("#selectCountType").val();
                    var selectCountTypeText = $("#selectCountType option:selected").text();
                    var selectYY = $("#selectYY").val();

                    var form = document.createElement('form');
                    var iframe = document.createElement('iframe');
                    var selectDBIn = document.createElement('input');
                    var selectCountTypeIn = document.createElement('input');
                    var selectCountTypeTextIn = document.createElement('input');
                    var selectYYIn = document.createElement('input');

                    selectDBIn.setAttribute("id", "selectDBIn");
                    selectDBIn.setAttribute("name", "selectDBIn");
                    selectDBIn.setAttribute("value", selectDB);

                    selectCountTypeIn.setAttribute("id", "selectCountTypeIn");
                    selectCountTypeIn.setAttribute("name", "selectCountTypeIn");
                    selectCountTypeIn.setAttribute("value", selectCountType);

                    selectCountTypeTextIn.setAttribute("id", "selectCountTypeTextIn");
                    selectCountTypeTextIn.setAttribute("name", "selectCountTypeTextIn");
                    selectCountTypeTextIn.setAttribute("value", selectCountTypeText);
                    
                    selectYYIn.setAttribute("id", "selectYYIn");
                    selectYYIn.setAttribute("name", "selectYYIn");
                    selectYYIn.setAttribute("value", selectYY);

                    iframe.setAttribute("name", "postiframe");
                    iframe.setAttribute("id", "postiframe");
                    iframe.setAttribute("style", "display: none");

                    form.appendChild(iframe);
                    form.appendChild(selectDBIn);
                    form.appendChild(selectCountTypeIn);
                    form.appendChild(selectCountTypeTextIn);
                    form.appendChild(selectYYIn);


                    $("body").append(form);

                    form.setAttribute("action", "handler/filedownloadlist.ashx");
                    form.setAttribute("method", "post");
                    form.setAttribute("enctype", "multipart/form-data");
                    form.setAttribute("encoding", "multipart/form-data");
                    form.setAttribute("target", "postiframe");
                    form.setAttribute("style", "display: none");
                    form.setAttribute("name", "form1");
                    form.submit();

                    intervalID = window.setInterval(KeepAskSessionStatus, 2000);
                });
            });

            function generateUUID() {
                var d = new Date().getTime();
                var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
                    var r = (d + Math.random() * 16) % 16 | 0;
                    d = Math.floor(d / 16);
                    return (c == 'x' ? r : (r & 0x3 | 0x8)).toString(16);
                });
                return uuid;
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

            function bindTable() {
                $.ajax({
                    type: "POST",
                    url: "handler/list.ashx",
                    async: false,
                    data: {
                        dbtype: $("#selectDB").val(),
                        CountType: $("#selectCountType").val(),
                        Year: $("#selectYY").val()
                    },
                    error: function (xhr, textStatus) {
                        console.log(xhr.responseText);
                        alert("查詢條件連繫錯誤");
                    },
                    //dataType: "text",
                    success: function (response) {
                        $("#genDIV").children().remove();

                        if (response.length > 0) {
                            var genhead = '<table width="60%" border="0" cellspacing="0" cellpadding="0" style="text-align:center;" align="center">';
                            genhead += '<tr>';
                            genhead += '<th>排名</th>';
                            genhead += '<th>單位</th>';
                            genhead += '<th>部門名稱</th>';
                            genhead += '<th>部門代碼</th>';
                            genhead += '<th>姓名</th>';
                            genhead += '<th>工號</th>';
                            genhead += '<th id="thCount">總使用次數</th></tr>';
                            var genbody = '';
                            var genend = '</table>';
                            for (var i = 0; i < response.length; i++) {
                                genbody += '<tr>';
                                genbody += '<td>' + parseInt(i + 1) + '</td>';
                                genbody += '<td>' + response[i].orgcname + '</td>';
                                genbody += '<td>' + response[i].main_deptcdcname + '</td>';
                                genbody += '<td>' + response[i].main_deptcd + '</td>';
                                genbody += '<td>' + response[i].cname + '</td>';
                                genbody += '<td>' + response[i].empno + '</td>';
                                genbody += '<td>' + response[i].SUNMt + '</td>';
                                genbody += '</tr>';
                            }

                            $("#genDIV").append(genhead + genbody + genend);

                        }
                        else {
                            $("#genDIV").append('<div style="text-align: center; padding-top:15px;" align="center">尚無資料</div>');
                        }
                    }
                });
            }
        </script>
</asp:Content>

