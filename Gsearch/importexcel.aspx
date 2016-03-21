<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="importexcel.aspx.cs" Inherits="importexcel" %>

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
                    <span class="tabmenuT2Btn"><a href="gsearch.aspx" target="_self">查詢條件</a></span>
                    <span class="tabmenuT2Btn tabmenuT2BtnCurrent"><a href="importexcel.aspx" target="_self">上傳檔案</a></span>
                </div>
                <!-- left -->
            </div>
            <!-- twocol -->
        </div>
        <!-- tabmenuT2block -->
        <!-- tabmenu -->
    <div class="fixwidth margin10TB gentable">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td colspan="4" style="text-align:center;">
                    每月1號為IP歷史清除日，請勿於此時上傳任何資料。
                </td>
            </tr>
  <tr>
    <th width="12%" align="right"><div class="titlebackicon">ACM</div></th>
    <td>
        <input type="file" class="genbtnS" name="FileACM" />
        <input id="btnACM" type="button" value="上傳" class="genbtnS Uploadbtn" name="ACM" />
        <input id="btnACMCancel" type="button" value="取消" class="genbtnS cancelbtn" />
    </td>
    <th align="right"><div class="titlebackicon">SciFinder</div></th>
    <td>
        <input class="genbtnS" name="FileSciFinder" type="file" />
        <input id="btnSciFinder" class="genbtnS Uploadbtn" name="SciFinder" type="button" value="上傳" />
        <input id="btnSciFinderCancel" class="genbtnS cancelbtn" type="button" value="取消" /> 
    </td>
  </tr>
  <tr>
    <th align="right"><div class="titlebackicon">ACS</div></th>
    <td>
        <input type="file" class="genbtnS" name="FileACS" />
        <input id="btnACS" type="button" value="上傳" class="genbtnS Uploadbtn" name="ACS" />
        <input id="btnACSCancel" type="button" value="取消" class="genbtnS cancelbtn" />
    </td>
    <th align="right"><div class="titlebackicon">SDOL</div></th>
    <td>
        <input class="genbtnS" name="FileSDOL" type="file" />
        <input id="btnSDOL" class="genbtnS Uploadbtn" name="SDOL" type="button" value="上傳" />
        <input id="btnSDOLCancel" class="genbtnS cancelbtn" type="button" value="取消" /> 
    </td>
  </tr>
  <tr>
    <th align="right"><div class="titlebackicon">ASTM</div></th>
    <td>
        <input type="file" class="genbtnS" name="FileASTM" />
        <input id="btnASTM" type="button" value="上傳" class="genbtnS Uploadbtn" name="ASTM" />
        <input id="btnASTMCancel" type="button" value="取消" class="genbtnS cancelbtn" />
    </td>
    <th align="right"><div class="titlebackicon">SPIE</div></th>
    <td>
        <input class="genbtnS" name="FileSPIE" type="file" />
        <input id="btnSPIE" class="genbtnS Uploadbtn" name="SPIE" type="button" value="上傳" />
        <input id="brnSPIECancel" class="genbtnS cancelbtn" type="button" value="取消" /> 
    </td>
  </tr>
  <tr>
    <th align="right"><div class="titlebackicon">Digitimes</div></th>
    <td>
        <input type="file" class="genbtnS" name="FileDigitimes" />
        <input id="btnDigitimes" type="button" value="上傳" class="genbtnS Uploadbtn" name="Digitimes" />
        <input id="btnDigitimesCancel" type="button" value="取消" class="genbtnS cancelbtn" />
    </td>
    <th align="right"><div class="titlebackicon">TI</div></th>
    <td>
        <input type="file" class="genbtnS" name="FileTI" />
        <input id="btnTI" type="button" value="上傳" class="genbtnS Uploadbtn" name="TI" />
        <input id="btnTICancel" type="button" value="取消" class="genbtnS cancelbtn" />
    </td>
  </tr>
  <tr>
    <th align="right"><div class="titlebackicon">EV</div></th>
    <td>
        <input class="genbtnS" name="FileEV" type="file" />
        <input id="btnEV" class="genbtnS Uploadbtn" name="EV" type="button" value="上傳" />
        <input id="btnEVCancel" class="genbtnS cancelbtn" type="button" value="取消" /> 

    </td>
    <th align="right"><div class="titlebackicon">Wiley</div></th>
    <td>
        <input type="file" class="genbtnS" name="FileWiley" />
        <input id="btnWiley" type="button" value="上傳" class="genbtnS Uploadbtn" name="Wiley" />
        <input id="btnWileyCancel" type="button" value="取消" class="genbtnS cancelbtn" />
    </td>
  </tr>
  <tr>
    <th align="right"><div class="titlebackicon">IHS Goldenfire</div></th>
    <td>
        <input class="genbtnS" name="FileIHS" type="file" />
        <input id="btnIHS" class="genbtnS Uploadbtn" name="Goldfire" type="button" value="上傳" />
        <input id="btnIHSCancel" class="genbtnS cancelbtn" type="button" value="取消" />
    </td>
    <th align="right"><div class="titlebackicon">工程</div></th>
    <td>
        <input type="file" class="genbtnS" name="工程" />
        <input id="btnEngineering" type="button" value="上傳" class="genbtnS Uploadbtn" name="工程" />
        <input id="btnEngineeringCancel" type="button" value="取消" class="genbtnS cancelbtn" />
    </td>
  </tr>
  <tr>
    <th align="right"><div class="titlebackicon">IEEE</div></th>
    <td>
        <input class="genbtnS" name="FileIEEE" type="file" />
        <input id="btnIEEE" class="genbtnS Uploadbtn" name="IEEE" type="button" value="上傳" />
        <input id="btnIEEECAncel" class="genbtnS cancelbtn" type="button" value="取消" />
    </td>
    <th align="right"><div class="titlebackicon">天下</div></th>
    <td>
        <input type="file" class="genbtnS" name="天下" />
        <input id="btnTangShang" type="button" value="上傳" class="genbtnS Uploadbtn" name="天下" />
        <input id="btnTangShangCancel" type="button" value="取消" class="genbtnS cancelbtn" />
    </td>
  </tr>
  <tr>
    <th align="right"><div class="titlebackicon">JCR</div></th>
    <td>
        <input class="genbtnS" name="FileJCR" type="file" />
        <input id="btnJCR" class="genbtnS Uploadbtn" name="JCR" type="button" value="上傳" />
        <input id="btnJCRCancel" class="genbtnS cancelbtn" type="button" value="取消" />
    </td>
    <th align="right"><div class="titlebackicon">萬方</div></th>
    <td>
        <input type="file" class="genbtnS" name="萬方" />
        <input id="btnWangFang" type="button" value="上傳" class="genbtnS Uploadbtn" name="萬方" />
        <input id="btnWangFangCancel" type="button" value="取消" class="genbtnS cancelbtn" />
    </td>
  </tr>
  <tr>
    <th align="right"><div class="titlebackicon">Nature</div></th>
    <td>
        <input class="genbtnS" name="Filenature" type="file" />
        <input id="btnnature" class="genbtnS Uploadbtn" name="nature" type="button" value="上傳" />
        <input id="btnnatureCancel" class="genbtnS cancelbtn" type="button" value="取消" /> 
    </td>
    <th align="right"><div class="titlebackicon">聯合知識庫</div></th>
    <td>
        <input class="genbtnS" name="聯合知識庫" type="file" />
        <input id="btnUnion" class="genbtnS Uploadbtn" name="聯合知識庫" type="button" value="上傳" />
        <input id="btnUnionCancel" class="genbtnS cancelbtn" type="button" value="取消" /> 
    </td>
  </tr>
  <tr>
    <th align="right"><div class="titlebackicon">Oxford</div></th>
    <td>
        <input class="genbtnS" name="FileOxford" type="file" />
        <input id="btnOxford" class="genbtnS Uploadbtn" name="Oxford" type="button" value="上傳" />
        <input id="btnOxfordCancel" class="genbtnS cancelbtn" type="button" value="取消" /> 
    </td>
    <th align="right"><div class="titlebackicon">Science</div></th>
    <td>
        <input class="genbtnS" name="FileScience" type="file" />
        <input id="btnScience" class="genbtnS Uploadbtn" name="Science" type="button" value="上傳" />
        <input id="btnScienceCancel" class="genbtnS cancelbtn" type="button" value="取消" /> 
    </td>
  </tr>
 
  <tr>
    <th align="right"><div class="titlebackicon">通用匯入</div></th>
    <td>
        <input class="genbtnS" name="FileUtility" type="file" />
        <input id="btnUtility" class="genbtnS Uploadbtn" name="通用" type="button" value="上傳" />
        <input id="btnUtilityCancel" class="genbtnS cancelbtn" type="button" value="取消" />
    </td>
    <th align="right">&nbsp;</th>
    <td>
        &nbsp;</td>
  </tr>
 
</table>
</div>
<%--dialog--%>
<div class="dialog ui-dialog-titlebar lineheight03">
    <div id="processbar" style="display:none;">
        <img src="images/loading.gif"/>
    </div>
    <div id="sendmsg"></div>

    <div id="chooseDate" style="display:none;">
        <div>
            年分:<input type="text" id="dialogYear" style="width: 40px;" /><br />
            月份:<input type="text" id="dialogMonth" style="width: 20px;" />
        </div>
    <div style="float:right;"><input type="button" value="送出" class="genbtnS" id="submitForm" /></div>
    </div>
</div>

    <script type="text/javascript">
        var UUID;
        var file;//選到的那個檔案
        var thisButton;//點擊的那個按鈕

        $("#dialogYear").val((new Date).getFullYear());
        $("#dialogMonth").val((new Date).getMonth());//起始是0

        $(document).ready(function () {
            //取消按鈕
            $(".cancelbtn").click(function () {
                var file = $(this).parent().children("input[type='file']");
                file.after(file.clone().val(""));
                file.remove();
            });

           
            //選擇年月
            $(".Uploadbtn").click(function () {
                file = $(this).parent().children("input[type='file']");

                if (file.length == 0 ) {
                    alert("請選擇檔案");
                    return false;
                }
                else if (file.length > 1) {
                    alert("一次請選擇一個檔案");
                    return false;
                }
                var re = /\.(xls|xlsx)$/i;  //允許的Excel副檔名
                if (!re.test(file[0].value)) {
                    alert("請選擇xls或xlsx檔案上傳");
                    return false;
                }

                if (isNaN($("#dialogYear").val())) {
                    alert("年份請輸入數字");
                    return false;
                }

                if (isNaN($("#dialogMonth").val())) {
                    alert("月份請輸入數字");
                    return false;
                }

                $("#sendmsg").text("");
                $(".dialog").dialog("open").dialog({ width: 175, height: 105 })
                $(".ui-dialog-titlebar").show();
                $(".dialog").dialog('option', 'title', "請選擇此檔案的日期");  
                $("#chooseDate").show();
                $("#processbar").hide();

                thisButton = $(this);
            });

            //上傳按鈕
            $("#submitForm").click(function () {
                var form = document.createElement('form');
                var iframe = document.createElement('iframe');
                var textType = document.createElement('input');
                var YearValinput = document.createElement('input');
                var MonthValinput = document.createElement('input');
                UUID = generateUUID();

                iframe.setAttribute("name", "postiframe");
                iframe.setAttribute("id", "postiframe");
                iframe.setAttribute("style", "display: none");

                textType.setAttribute("id", "textType");
                textType.setAttribute("name", "textType");
                textType.setAttribute("value", thisButton.attr("name"));

                YearValinput.setAttribute("id", "YearVal");
                YearValinput.setAttribute("name", "YearVal");
                YearValinput.setAttribute("value", $("#dialogYear").val());
                
                MonthValinput.setAttribute("id", "MonthVal");
                MonthValinput.setAttribute("name", "MonthVal");
                MonthValinput.setAttribute("value", $("#dialogMonth").val());
                
                form.appendChild(iframe);
                form.appendChild(file.get(0));
                form.appendChild(textType);
                form.appendChild(YearValinput);
                form.appendChild(MonthValinput);
                $("body").append(form);

                form.setAttribute("action", "handler/importexcel.ashx");
                form.setAttribute("method", "post");
                form.setAttribute("enctype", "multipart/form-data");
                form.setAttribute("encoding", "multipart/form-data");
                form.setAttribute("target", "postiframe");
                form.setAttribute("style", "display: none");
                form.setAttribute("name", UUID);
                form.submit();

                
                //file不見了 需要再補回去
                $(".ui-dialog-titlebar").hide();
                thisButton.before(file.clone().val(""));
                $(".dialog").dialog('option', 'width', 240).dialog('option', 'height', 70);
                $("#sendmsg").text("上傳中請稍後...");
                $("#processbar").show();
                $("#chooseDate").hide();     
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

        function ajaxAlertFun(msg) {
            //form submit出去後,即可清除
            $("form[name='" + UUID + "']").remove();
            $(".ui-dialog-titlebar").show();
            $("#processbar").hide();
            $(".dialog").dialog({ width: 240, height: 'auto' });

            if (msg.indexOf('error') > -1) {
                //error_錯誤訊息
                //$(".dialog").dialog("open");
                $(".dialog").dialog('option', 'title', "警告");
                $("#sendmsg").text(msg.substring(msg.indexOf('_') + 1, msg.length));
            }
            else {
                $(".dialog").dialog('option', 'title', "檔案上傳");
                $("#sendmsg").text("上傳成功!!");
            }
            
        }
    </script>
</asp:Content>

