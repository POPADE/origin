<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%
    String basePath = request.getScheme() + "://" + request.getServerName()
            + ":" + request.getServerPort()  + request.getContextPath() +"/";
%>
<html>
<head>
    <base href="<%=basePath%>">
    <!--JQUERY-->
    <script type="text/javascript" src="jquery/jquery-1.11.1-min.js"></script>
    <!--BOOTSTRAP框架-->
    <link rel="stylesheet" href="jquery/bootstrap_3.3.0/css/bootstrap.min.css">
    <script type="text/javascript" src="jquery/bootstrap_3.3.0/js/bootstrap.min.js"></script>
    <!--BOOTSTRAP_DATETIMEPICKER插件-->
    <link rel="stylesheet" href="jquery/bootstrap-datetimepicker-master/css/bootstrap-datetimepicker.min.css">
    <script type="text/javascript" src="jquery/bootstrap-datetimepicker-master/js/bootstrap-datetimepicker.js"></script>
    <script type="text/javascript" src="jquery/bootstrap-datetimepicker-master/locale/bootstrap-datetimepicker.zh-CN.js"></script>
    <title>演示bs_datetimepicker插件</title>

<script type="text/javascript">
    $(function (){
        //当容器加载完成，对容器调用工具函数
        $("#myDate").datetimepicker({
            language:'zh-CN',
            format:'yyyy-mm-dd',
            minView:'month',
            initialData:new Date(),
            autoclose:true,
            todayBtn:true,
            clearBtn:true
        });
    })
</script>
</head>
<body>
<input type="text" id="myDate" readonly>

</body>
</html>
