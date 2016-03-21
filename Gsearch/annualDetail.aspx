<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="annualDetail.aspx.cs" Inherits="annualDetail" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<input id="btnBack" type="button" value="返回上頁" class="genbtn searchbtn delBtn" onclick="javascript:history.go(-1);" />

<div style="text-align: center; padding-top:15px;" align="center">
    <div id="ShowTitle"></div>
    選擇類型：<select id="selectCountType">
            <option value="">全部</option>
            <option value="main_downloadC">下載數</option>
            <option value="main_viewC">瀏覽數</option>
            <option value="main_indexC">檢索數</option>
            <option value="main_logC">登入數</option>
             </select>
    <input type="button" value="匯出" class="genbtn searchbtn" id="ExportAllbtn" />
</div>
        <div class="stripeMe margin15TB" style="padding:0 10px;">
        <div id="genDIV"></div>
    </div>

<script type="text/javascript">
    var intervalID;

    $(document).ready(function () {
        bindTable();

        $("#selectCountType").change(function () {
            bindTable();
            //if ($("#selectCountType").val() != '') {
            //    $("#thCount").text($("#selectCountType :selected").text());
            //}
        });


        $(document).on("click", "#ExportAllbtn", function (event) {
            $(".dialog").dialog("open").dialog({ width: 240, height: 70 })
            $("#sendmsg").text("下載中請稍後...");
            $("#processbar").show();

            var uid = $.getParamValue('uid');
            var selectCountType = $("#selectCountType option:selected").val();
            var selectCountTypeText = $("#selectCountType option:selected").text();

            var form = document.createElement('form');
            var iframe = document.createElement('iframe');
            var uidIn = document.createElement('input');
            var selectCountTypeIn = document.createElement('input');
            var selectCountTypeTextIn = document.createElement('input');

            uidIn.setAttribute("id", "uid");
            uidIn.setAttribute("name", "uid");
            uidIn.setAttribute("value", uid);

            selectCountTypeIn.setAttribute("id", "selectCountType");
            selectCountTypeIn.setAttribute("name", "selectCountType");
            selectCountTypeIn.setAttribute("value", selectCountType);

            selectCountTypeTextIn.setAttribute("id", "selectCountTypeText");
            selectCountTypeTextIn.setAttribute("name", "selectCountTypeText");
            selectCountTypeTextIn.setAttribute("value", selectCountTypeText);

            iframe.setAttribute("name", "postiframe");
            iframe.setAttribute("id", "postiframe");
            iframe.setAttribute("style", "display: none");

            form.appendChild(iframe);
            form.appendChild(uidIn);
            form.appendChild(selectCountTypeIn);
            form.appendChild(selectCountTypeTextIn);

            $("body").append(form);

            form.setAttribute("action", "handler/filedownloadAnnual.ashx");
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

    //$.getParamValue('gid')
    //------取得URL的get參數 做ajax 把資料撈出來------------//
    $.extend({
        getParamValue: function (paramName) {
            parName = paramName.replace(/[\[]/, '\\\[').replace(/[\]]/, '\\\]');
            var pattern = '[\\?&]' + paramName + '=([^&#]*)';
            var regex = new RegExp(pattern);
            var matches = regex.exec(window.location.href);
            if (matches == null) { return ''; }
            else { return decodeURIComponent(matches[1].replace(/\+/g, ' ')); }
        },
        ResizeDialog: function () {
            var dialogwidth = $("#dialog").outerWidth();
            var Screenwidth = $(document).outerWidth();

            var dialogheight = $("#dialog").outerHeight();
            var Screenheight = $(document).outerHeight();

            var x = (Screenwidth - dialogwidth) / 2;
            var y = (Screenheight - dialogheight) / 2 - 100;
            $("#dialog").dialog('option', 'position', [x, y]);
        }
    });

    function bindTable() {
        $.ajax({
            type: "POST",
            url: "handler/annualDetail.ashx",
            async: false,
            data: {
                request: $.getParamValue('uid'),
                CountType: $("#selectCountType").val()
            },
            error: function (xhr, textStatus) {
                console.log(xhr.responseText);
                alert("查詢條件連繫錯誤");
            },
            //dataType: "text",
            success: function (response) {
                $("#genDIV").children().remove();

                if (response.length > 0) {
                    var genhead = '<table width="100%" border="0" cellspacing="0" cellpadding="0" style="text-align:center;">';
                    genhead += '<tr>';
                    genhead += '<th>排名</th>';
                    genhead += '<th>單位代碼</th>';
                    genhead += '<th>單位名稱</th>';
                    genhead += '<th>百分比</th>';
                    genhead += '<th>1月</th>';
                    genhead += '<th>2月</th>';
                    genhead += '<th>3月</th>';
                    genhead += '<th>4月</th>';
                    genhead += '<th>5月</th>';
                    genhead += '<th>6月</th>';
                    genhead += '<th>7月</th>';
                    genhead += '<th>8月</th>';
                    genhead += '<th>9月</th>';
                    genhead += '<th>10月</th>';
                    genhead += '<th>11月</th>';
                    genhead += '<th>12月</th>';
                    genhead += '<th>小計</th>';
                    genhead += '</tr>';
                    var genbody = '';
                    var genend = '</table>';
                    for (var i = 0; i < response.length; i++) {
                        genbody += '<tr>';
                        genbody += '<td>' +( i + 1 )+ '</td>';
                        genbody += '<td>' + response[i].main_org + '</td>';
                        genbody += '<td>' + response[i].main_orgcname + '</td>';
                        genbody += '<td>' + response[i].main_percentage + '</td>';
                        genbody += '<td>' + response[i].Jan + '</td>';
                        genbody += '<td>' + response[i].Feb + '</td>';
                        genbody += '<td>' + response[i].Mrz + '</td>';
                        genbody += '<td>' + response[i].Apr + '</td>';
                        genbody += '<td>' + response[i].May + '</td>';
                        genbody += '<td>' + response[i].Jun + '</td>';
                        genbody += '<td>' + response[i].Jul + '</td>';
                        genbody += '<td>' + response[i].Aug + '</td>';
                        genbody += '<td>' + response[i].Sep + '</td>';
                        genbody += '<td>' + response[i].Oct + '</td>';
                        genbody += '<td>' + response[i].Nov + '</td>';
                        genbody += '<td>' + response[i].Dez + '</td>';
                        genbody += '<td>' + response[i].total + '</td>';
                        genbody += '</tr>';
                    }

                    $("#genDIV").append(genhead + genbody + genend);
                    var titledbtype = response[0].dbtype==''?'全部':response[0].dbtype;
                    $("#ShowTitle").html("各單位於<span class='font-title font-size3'>" + titledbtype + "</span>資料庫在<span class='font-title font-size3'>" + response[0].year + "</span>年使用量<br />");
                }
                else {
                    $("#genDIV").append('<div style="text-align: center; padding-top:15px;" align="center">尚無資料</div>');
                }
            }
        });
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
        </script>
</asp:Content>

