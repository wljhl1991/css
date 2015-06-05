 <%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
    <%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c" %> 
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>楼宇党建工作</title>
<script src="js/jquery-1.7.1.min.js" type="text/javascript"></script>
<script src="js/jquery.uploadify.min.js" type="text/javascript"></script>
<script src="js/lydj.js" type="text/javascript"></script>
<link href="css/uploadify.css" rel="stylesheet" type="text/css" >
<link href="css/button.css" rel="stylesheet" type="text/css" />
</head>

<body>
<center>
	<h1>楼宇党建文档汇总</h1>

	采集模版下载:
	<br/>
	<form action="downloadT.do" method="post" >
	<input type="hidden" id ="streetInfo" name="streetInfo" value=""></input>
	<input type="hidden" id ="localFolad" name="localFolad" value="${ localFolad}"></input>
	1.区县:<select id = "town" name = "town" onchange="change();">
	<option value ="-"  selected="selected" > 请选择</option>
		<c:forEach var="t" items="${town }" >
		  <%--<c:if test="${t.key =='110105000'}"> selected="selected" </c:if>   <c:if test="${t.key =='110105000'}"><option value ="${t.key }" selected="selected">${t.value }</option> </c:if> --%>
		<option value ="${t.key }"  > ${t.value }</option>
		</c:forEach>
	</select>
	2.街道:
 <span style="position:absolute;border:1pt solid #c1c1c1;overflow:hidden;width:188px;height:19px;clip:rect(-1px 190px 190px 170px);"> 
	<!-- <select id = "street" name ="street" > -->
 	<select id = "street" name ="street" style="width:190px;height:20px;margin:-2px;" onChange="javascript:document.getElementById('ccdd').value=document.getElementById('street').options[document.getElementById('street').selectedIndex].text;"> 
	<c:forEach var="s" items="${street }" >
		<option value ="${s.key }">${s.value }</option>
		</c:forEach>
	</select>
 	</span> 
	 <span style="position:absolute;border-top:1pt solid #c1c1c1;border-left:1pt solid #c1c1c1;border-bottom:1pt solid #c1c1c1;width:170px;height:19px;"> 
		<input type="text" name="ccdd" id="ccdd" value="选择或输入园区" style="width:170px;height:15px;border:0pt;" onkeyup="changInput();" onfocus="clickInput();" onblur="cleanSelect();"> 
	<div id ="showContane" style="margin-top:-12px;overflow:auto;width:170px;border: 1px solid #ccc;">
		 <!-- <ul style="display: 'none'">
		<li >ddddddddddd</li>
		</ul>  -->
	</div>
	</span> 
 <br/>
	<!-- onclick="downloadtemp('templet');" -->
	3.楼宇名称:<input type="text" name = "name" id="name" value="输入楼宇名称" onfocus="cleanInput();" onblur="insertInput();"></input>
	<input id='downloadTmp'  class="btn12" type="submit" value="下载填写表格"
		onmouseout="this.style.backgroundPosition='left top'"
		onmouseover="this.style.backgroundPosition='left -40px'"
		style="background-position: left top;">
		
		<p></p>
		<br/>
		<br/>
		<br/>
		<a href="notic.jsp" target="blank"> 遇到问题？点击此处</a><br>
		<hr>
		汇总：
		<div id="queue"></div>
		<input id="file_upload" name="file_upload" type="file" multiple="true"><div id ="count" value=""> 所选文件数量：0</div>
		<br>     
<!--    <input type="checkbox" name="checkbox" value="checkbox1"/>(忽略工作站名称相同，方便测试使用！) -->
		<input name ="typeTodo" type="radio" value="start.do" checked="checked"/>街道（乡镇）汇总
		<br>
		<input name ="typeTodo" type="radio" value="start2.do" />区（县）/市委汇总
		</form>
	<div id = "downloadto"></div>
	<div id = "printto"></div>
		 <button id="uploadAndGather" onclick="uploadAndGather();" class="btn29" 
				onmouseover="this.style.backgroundPosition='left -43px'"
				onmouseout="this.style.backgroundPosition='left top'" >上传并汇总</button>
				<br/>
	<!--  <button id="stopUploadAndGather" onclick="stopUploadAndGather();"> 停止上传</button>  -->
<!-- <p><a href="javascript:$('#file_upload').uploadify('upload','*')" >上传并汇总</a></p>
<p><a href="javascript:$('#file_upload').uploadify('stop')">Stop the Uploads!</a></p> -->
	</center>
	
	
	
</body>

</html>