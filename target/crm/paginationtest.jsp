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

    <!--BOOTSTRAP-->
    <link rel="stylesheet" type="text/css" href="jquery/bootstrap_3.3.0/css/bootstrap.min.css">
    <script type="text/javascript" src="jquery/bootstrap_3.3.0/js/bootstrap.min.js"></script>

    <!--PAGINATION plugin-->
    <link rel="stylesheet" type="text/css" href="jquery/bs_pagination-master/css/jquery.bs_pagination.min.css">
    <script type="text/javascript" src="jquery/bs_pagination-master/js/jquery.bs_pagination.min.js"></script>
    <script type="text/javascript" src="jquery/bs_pagination-master/localization/en.js"></script>
    <title>演示bs_pagination插件的使用</title>

    <script type="text/javascript">

        $(function (){
            $("#demo_pag1").bs_pagination({
                currentPage: 1, //当前页号 相当于pageNo
                rowsPerPage: 10, //每页显示条数 相当于pageSize
                totalRows: 1000, //总条数

                visiblePageLinks: 5, //最多可以显示的卡片数

                showGoToPage: true,  //是否显示跳转页数
                showRowsPerPage: true, //是否显示每页显示条数
                showRowsInfo: true, //是否显示记录的信息

                totalPages: 100,//总页数 必填参数

                //用户每次切换页号，都自动触发本函数
                //每次返回切换页号之后的pageNo pageSize
                onChangePage: function (event, pageObj) {
                    alert(pageObj.currentPage);
                    alert(pageObj.rowsPerPage);
                }
            })
        })

    </script>
</head>
<body>
<div id="demo_pag1"></div>
</body>
</html>
