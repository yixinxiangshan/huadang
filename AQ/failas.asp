<%
'初始化基本信息
session("M_DurationTime") = 0    '持续时间
session("M_RightNum") = 0        '答对的题目数
session("M_ErrorNum") = 0        '答错的题目数
session("M_StarTime") = 0        '开始时间
session("M_EndTime") = 0         '结束时间
session("lastRequestTimer") = 0  '最后读取的时间
session("ydQuestions") = ""      '已经答过的题目
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>隆重纪念中国共产党成立90周年--党史知识竞赛</title>
<link href="css.css" rel="stylesheet" type="text/css">
</head>
<body>
<div align="center">
<table width="960" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
    <!--#include file= include/top.asp-->
    </td>
  </tr>
  <tr>
    <td height="400" align="center" valign="top" bgcolor="#FFFFFF"><p><img src="images/Fail_Info.jpg" width="477" height="365" border="0" usemap="#Map1"></p></td>
  </tr>
  <tr>
    <td height="15" align="center"></td>
  </tr>
  <tr>
    <td>
     <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" valign="top" class="fontYellowHeight23">网站简介 | <a href="/hhdj/webpage/comment.asp" target="_blank" class="yellowNagative">互动交流</a> | <a href="/hhdj/bbs" class="yellowNagative">在线论坛</a> | <a href="mailto:shuji@huahong.com.cn" class="yellowNagative">书记信箱</a> | <a href="/hhdj/webpage/sitemap.htm" target="_blank" class="yellowNagative">网站地图</a> | <a href="/hhdj/webpage/copyright.htm" target="_blank" class="yellowNagative">版权声明</a><br>
          2011 中共上海华虹（集团）有限公司委员会主办 版权所有 <br>
          最佳浏览 1024x768 分辨率 </td>
        <td align="right" valign="top" class="fontYellowHeight23">未经授权禁止转载、摘编、复制或建立镜像<br>
          Power by Vasisoft</td>
      </tr>
     </table>
    </td>
  </tr>
</table>
</div>

<map name="Map1">
  <area shape="rect" coords="270,124,336,139" href="index.asp" target="_self">
</map>

</body>
</html>
