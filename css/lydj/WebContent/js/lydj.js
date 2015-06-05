		$(function() {
			$('#file_upload').uploadify({
				'auto'     : false,
				'buttonText' : '批量添加 文件',
				/*'formData' : {'localFolad':'${ localFolad}'},*/
				'formData' : {'localFolad': $('#localFolad').val()},
				'removeCompleted' : false,
				'swf'      : 'image/uploadify.swf',
				'uploader' : 'upload2.do',
				'multi'    : true ,
				'fileTypeExts' : '*.xlsx;*.zip',
				'uploadLimit' : 500,
				'onSelect' : function(file){
					setTimeout('munShow()',200);
				} ,				
				'onQueueComplete' : function(queueData) {
					var todo  = $("input[name='typeTodo']:checked").val();
					var testok  = $("input[name='checkbox']").attr("checked");
					var testStyle  = $("input[name='copyStyle']").attr("checked");
					$.ajaxSetup({
						  contentType: "application/x-www-form-urlencoded; charset=utf-8"
						});
					$.post(todo,{'localFolad':$('#localFolad').val(),'testok':testok,'testStyle' :testStyle},function(data){
						if(data.err == "错误"){
  							alert((data.result).replace(/<br\/>/g,""));  
							 $('#downloadto').html(data.result); 
						}else{
		            		alert("数据汇总已完成!");
		            	 $('#downloadto').html("  <button id='downloadShow' onclick='downloadthis(\""+data.result+"\");'  class=\"btn26\" value=\"Submit\" onmouseover=\"this.style.backgroundPosition='left -36px'\" onmouseout=\"this.style.backgroundPosition='left top'\" >下载电子版</button>"); 
		            	 /* $('#printto').html("  <button id='downloadPrint' onclick='printthis(\""+data.result+"\");'  class=\"btn26\" value=\"Submit\" onmouseover=\"this.style.backgroundPosition='left -36px'\" onmouseout=\"this.style.backgroundPosition='left top'\" >下载打印版</button>");  */
						}
					});
					alert("上传已完成");
		        },
		        'onCancel' : function(file) {
		        	var size = $('#file_upload-queue').children().size();
		        	$('#count').attr("value",Number(Number(size)-1));
					$('#count').html("所选文件数量："+Number(Number(size)-1)); 
					/* setTimeout(opButton(),6000); */
					opButton(file.name);
		        }
			});
		});
		
		function downloadthis(result){
		 	window.open("download/"+result+".do"); 
		};
		function downloadtemp(result){
			var town = $('#town').find("option:selected").val();
			var street = $('#street').find("option:selected").val();
			var data = result+"_"+town+"_"+street;
		 	 window.open("download/"+data+".do"); 
		};
		function printthis(result){
		 	window.open("download/downPrint"+result+".do"); 
	 	/* $.get("download/"+result+".do",null);  */
		};
		function uploadAndGather(){
			var size = $('#file_upload-queue').children().size();
			size =Number(Number(size)-1);
			if(size >= 15){
				var a=confirm(size +"文件数量较多，若上传后没有反应，可将文件压缩为zip文件进行上传");
				if(a == false){
					$('#file_upload').uploadify('upload','*');
				}
			}else{
				$('#file_upload').uploadify('upload','*');
			}
		};
		/* function stopUploadAndGather(){
			$('#file_upload').uploadify('stop');
		}; */
		
		function checkFileName(fileName) {
			var g = false;
			fileName = fileName.substring(0, fileName.lastIndexOf("."));
			$("span.fileName")
					.each(
							function() {
								var fn = $(this).text();
								fn = fn.substring(0, fn.lastIndexOf("."));
								if (fileName.indexOf(fn) === -1) {
									/* var pattern = new RegExp("[`~!@#$^&*()=|{}':;',\\[\\].<>/?~！@#￥……&*（）&mdash;—|{}【】‘；：”“'。，、？]")  */
									var pattern = new RegExp(
											"[`~!@#$^&*%|{}'',\\[\\].<>/?@#￥……&*&mdash;|‘”“'？]")
									/* if(fn.indexOf("'")!=-1 ||fn.indexOf("\"")!=-1 ||fn.indexOf("%")!=-1 ||fn.indexOf("//")!=-1 ){
										
									} */
									
									 /* if (pattern.test(fn)) {
										g = true;
										if (!($(this).next().text() == "含有特殊字符")) {
											$(this)
													.after(
															"<font color='red'>含有特殊字符</font>");
											 //$('#uploadAndGather').attr('disabled',true);   
										}
									}  */
									
								}
							});
			return g;
		};
		function opButton(fileName) {
			$('#uploadAndGather').attr('disabled', checkFileName(fileName));
		};
		function munShow() {
			var size = $('#file_upload-queue').children().size();
			$('#count').html("");
			$('#count').attr("value", "");
			$('#count').attr("value", size);
			$('#count').html("所选文件数量：" + size);
			opButton("");
		}
		var thisid;
		function change(){
			var id = $('#town').val();
			thisis = id;
			 $.getJSON('changeData.do',{'id':id},function(data){
				var street = $('#street');
				street.empty();
				var options ="";
				$.each(data, function(i, field){
					options+="<option value="+i+">"+field+"</option>"; 
				});
					street.html(options);
			}); 
			$('#ccdd').val("");
			$('#streetInfo').val("");
			 changInput();
		}
		function test(){

		};
		function changInput(){
			var options ="";
			var town = $('#town').find("option:selected").val();
			var ccdd = $('#ccdd').val();
			 $('#street').find("option:selected").val("");
			$('#streetInfo').val(ccdd);
			 $.getJSON('changeData.do',{'id':town},function(data){
					$.each(data, function(i, field){
						if(""==ccdd  ){
							options+="<li onclick=\"selectItem('"+i+"','"+field+"');\">"+field+"</li>"; 
						}else if(field.indexOf(ccdd)!=-1){
						options+="<li onclick=\"selectItem('"+i+"','"+field+"');\">"+field+"</li>"; 
						}
/* 						options+="<option value="+i+">"+field+"</option>";  */
					});
						/* street.html(options); */
			var s = "<ul class='showSelect'>"+options+"</ul>";
			
			$('#showContane').html(s);
				}); 
		}
		
		
		
		function selectItem(code,cap){
			var street = $('#street').find("option:selected").val(code);
			var ccdd = $('#ccdd').val(cap);
			$('#showContane').css("display","none");
		}
		function clickInput(){
			$('#showContane').css("display","");
			if("选择或输入园区" ==  $('#ccdd').val()){
				$('#ccdd').val("");
			}
			changInput();
		}
		
		function cleanSelect(){
			$('#showContane').css("display","none");
			if("" ==  $('#ccdd').val().trim()){
				$('#ccdd').val("选择或输入园区");
			}
			$('#streetInfo').val($('#ccdd').val());				
		}
		function cleanInput(){
			if("输入楼宇名称" ==  $('#name').val()){
				$('#name').val("");
			}
		}
		function insertInput(){
			if("" ==  $('#name').val().trim()){
				$('#name').val("输入楼宇名称");
			}
		}