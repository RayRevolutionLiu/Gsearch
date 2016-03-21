<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="index.aspx.cs" Inherits="index" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
     <!-- tabmenu -->
 <div class="tabmenuT2block">
    <div class="fixwidth twocol">
        <div class="textcenter">
            <!-- 選單位置:left:置左、right:置右、center:置中 -->
            <span class="tabmenuT2Btn tabmenuT2BtnCurrent"><a href="index.aspx" target="_self">年度報表</a></span>
            <span class="tabmenuT2Btn"><a href="list.aspx" target="_self">最常使用者</a></span>
            <span class="tabmenuT2Btn"><a href="gsearch.aspx" target="_self">查詢條件</a></span>
            <span class="tabmenuT2Btn"><a href="importexcel.aspx" target="_self">上傳檔案</a></span>
        </div>
        <!-- left -->
    </div><!-- twocol -->
</div><!-- tabmenuT2block -->    
    <!-- tabmenu -->
    <div style="text-align: center; padding-top:15px;" align="center">
        <span>
        選擇年分：<select id="selectYY"></select></span>
       <span style="float:right;"><input type="button" value="刪除此年份所有資料" id="delAnnualYear" class="genbtn searchbtn delBtn" /></span>
    </div>
    <div class="stripeMe margin15TB" style="padding:0 10px;">
        <div id="genDIV"></div>
    	<%--<table width="100%" border="0" cellspacing="0" cellpadding="0" style="text-align:center;">
  <tr>
    <th></th>
    <th>1月</th>
    <th>2月</th>
    <th>3月</th>
    <th>4月</th>
    <th>5月</th>
    <th>6月</th>
    <th>7月</th>
    <th>8月</th>
    <th>9月</th>
    <th>10月</th>
    <th>11月</th>
    <th>12月</th>
  </tr>
  <tr>
    <th width="5%">ACM</th>
    <td><a href="#">12</a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>ACS</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="#">14</a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>ASTM</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="#">12256</a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="#">76567</a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>Digitimes</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="#">684908</a></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>EV</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>Goldfire</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>IEEE</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="#">1265</a></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>JCR</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="#">76552</a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>Nature</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="#">87432</a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>OUP</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="#">4356</a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>Science</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="#">2542</a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>SciFinder</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="#">5543</a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>SDOL</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="#">768532</a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>SPIE</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>SCOPUS</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>TI</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>Wiley</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>工程</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>天下</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>萬方</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th>聯合知識庫</th>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>--%>
    </div>
    <div>註:表格內數字為匯入的列數，不等於資料庫內的實際筆數</div>
    <%--dialog--%>
    <div class="dialog ui-dialog-titlebar lineheight03">
        <div id="processbar" style="display: none;">
            <img src="images/loading.gif" />
        </div>
        <div id="sendmsg"></div>
    </div>

    <script type="text/javascript">
        $(document).ready(function () {
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
                    var optonSTR = '';
                    if (response.length > 0) {
                        for (var i = 0; i < response.length; i++) {
                            optonSTR += "<option value='" + response[i].value + "'>" + response[i].cname + "</option>";
                        }
                    }
                    $("#selectYY").append(optonSTR);
                }
            });

            bindTable();


            $("#selectYY").change(function () {
                bindTable();
            });

            $(document).on("click", ".delBtn", function (event) {
                var year = $("#selectYY").val();
                if (year == '') {
                    alert("尚無資料");
                    return false;
                }
                else {
                    var conf;
                    var uid = $(this).attr("uid");
                    if (uid == '' || typeof (uid) == "undefined") {
                        uid = '';
                        conf = confirm('確定要刪除' + year + '年的所有資訊嗎?');
                    }
                    else {
                        var headRow = $(this).parents("tr").find("td:first").text();//資料庫名稱
                        var cellAndRow = $(this).parents('td,tr');
                        var cellIndex = cellAndRow[0].cellIndex;//被按下的按鈕在第幾欄
                        var headCol = $(this).parents("table").find("tr:first").find("th:eq(" + cellIndex + ")").text();//第幾月
                        conf = confirm('確定要刪除' + year + '年' + headRow + '資料庫在' + headCol + '的所有資訊嗎?');
                    }
         

                    if (conf) {
                        $.ajax({
                            type: "POST",
                            url: "handler/deleteAnnualData.ashx",
                            data: {
                                year: year,
                                uid: uid
                            },
                            error: function (xhr, textStatus) {
                                console.log(xhr.responseText);
                                alert("刪除資料錯誤");
                            },
                            //dataType: "text",
                            success: function (response) {
                                alert("刪除成功!");
                                bindTable();
                            },
                            beforeSend: function () {
                                $(".dialog").dialog("open").dialog({ width: 240, height: 70 })
                                $("#sendmsg").text("刪除中請稍後...");
                                $("#processbar").show();
                            },
                            complete: function () {
                                $(".dialog").dialog("close");
                            }
                        });
                    }
                }
            });
        });

        function bindTable() {
            $.ajax({
                type: "POST",
                url: "handler/listAnnualTable.ashx",
                async: false,
                data: {
                    dbtype: $("#selectYY").val()
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
                        genhead += '<th></th>';
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
                        for (var i = 0; i < response[0].NumberTable.length; i++) {
                            genbody += '<tr>';
                            genbody += '<td>' + response[0].NumberTable[i].statistics_dbtype + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].Jan == '' ? '' : response[0].NumberTable[i].Jan + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].Jan + '">刪除</a>') + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].Feb == '' ? '' : response[0].NumberTable[i].Feb + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].Feb + '">刪除</a>') + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].Mrz == '' ? '' : response[0].NumberTable[i].Mrz + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].Mrz + '">刪除</a>') + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].Apr == '' ? '' : response[0].NumberTable[i].Apr + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].Apr + '">刪除</a>') + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].May == '' ? '' : response[0].NumberTable[i].May + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].May + '">刪除</a>') + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].Jun == '' ? '' : response[0].NumberTable[i].Jun + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].Jun + '">刪除</a>') + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].Jul == '' ? '' : response[0].NumberTable[i].Jul + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].Jul + '">刪除</a>') + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].Aug == '' ? '' : response[0].NumberTable[i].Aug + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].Aug + '">刪除</a>') + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].Sep == '' ? '' : response[0].NumberTable[i].Sep + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].Sep + '">刪除</a>') + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].Oct == '' ? '' : response[0].NumberTable[i].Oct + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].Oct + '">刪除</a>') + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].Nov == '' ? '' : response[0].NumberTable[i].Nov + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].Nov + '">刪除</a>') + '</td>';
                            genbody += '<td>' + (response[0].NumberTable[i].Dez == '' ? '' : response[0].NumberTable[i].Dez + '<a href="javascript:void(0);" class="font-size1 padding5RL delBtn" uid="' + response[0].UidTable[i].Dez + '">刪除</a>') + '</td>';
                            genbody += '<td><a href="annualDetail.aspx?uid=' + response[0].UidTable[i].total + '">' + response[0].NumberTable[i].total + '</a></td>';
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

