package com.bjpowernode.crm.workbench.web.controller;

import com.bjpowernode.crm.commons.constants.Constants;
import com.bjpowernode.crm.commons.domain.ReturnObject;
import com.bjpowernode.crm.commons.utils.DateUtils;
import com.bjpowernode.crm.commons.utils.HSSFUtils;
import com.bjpowernode.crm.commons.utils.UUIDUtils;
import com.bjpowernode.crm.settings.domain.User;
import com.bjpowernode.crm.settings.service.UserService;
import com.bjpowernode.crm.workbench.domain.Activity;
import com.bjpowernode.crm.workbench.domain.ActivityRemark;
import com.bjpowernode.crm.workbench.service.ActivityRemarkService;
import com.bjpowernode.crm.workbench.service.ActivityService;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;


import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.*;
import java.util.*;


@Controller
public class ActivityController {

    @Autowired
    private UserService userService;

    @Autowired
    private ActivityService activityService;

    @Autowired
    private ActivityRemarkService activityRemarkService;

    @RequestMapping("/workbench/activity/index.do")
    public String index(HttpServletRequest request) {
        List<User> userList = userService.queryAllUsers();
        request.setAttribute("userList",userList);
        return "workbench/activity/index";

    }

    @RequestMapping("/workbench/activity/saveCreateActivity.do")
    @ResponseBody
    public Object saveCreateActivity(Activity activity, HttpSession session) {
        User user = (User) session.getAttribute(Constants.SESSION_USER);
        //封装参数
        activity.setId(UUIDUtils.getUUID());
        activity.setCreateTime(DateUtils.formatDateTime(new Date()));
        activity.setCreateBy(user.getId());

        ReturnObject returnObject = new ReturnObject();
        try {
            //调用service层，保存创建的市场活动
            int result = activityService.saveCreateActivity(activity);

            if (result > 0) {
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_SUCCESS);
            } else {
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
                returnObject.setMessage("系统忙，请稍后重试...");
            }
        } catch (Exception e) {
            e.printStackTrace();
            returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
            returnObject.setMessage("系统忙，请稍后重试...");
        }

        return returnObject;
    }

    @RequestMapping("/workbench/activity/queryActivityByConditionForPage.do")
    @ResponseBody
    public Object queryActivityByConditionForPage(String name, String owner, String startDate,
                                                  String endDate, Integer pageNo, Integer pageSize){
        //封装参数
        Map<String, Object> map = new HashMap<>();
        map.put("name", name);
        map.put("owner", owner);
        map.put("startDate", startDate);
        map.put("endDate", endDate);
        map.put("beginNo", (pageNo-1)*pageSize);
        map.put("pageSize", pageSize);
        //调用service方法，查询数据
        List<Activity> activityList = activityService.queryActivityByConditionForPage(map);
        int totalRows = activityService.queryCountOfActivityByCondition(map);

        //根据查询结果，生成响应信息
        Map<String, Object> resMap = new HashMap<>();
        resMap.put("activityList", activityList);
        resMap.put("totalRows", totalRows);
        return resMap;
    }

    @RequestMapping("/workbench/activity/deleteActivityIds.do")
    @ResponseBody
    public Object deleteActivityIds(String[] id){
        ReturnObject returnObject = new ReturnObject();
        try {
            //调用service层方法，删除市场活动
            int res = activityService.deleteActivityByIds(id);
            if(res > 0) {
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_SUCCESS);
            } else {
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
                returnObject.setMessage("系统忙，请稍后重试...");
            }
        } catch (Exception e) {
            e.printStackTrace();
            returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
            returnObject.setMessage("系统忙，请稍后重试...");
        }
        return returnObject;
    }

    @RequestMapping("/workbench/activity/queryActivityById.do")
    @ResponseBody
    public Object queryActivityById(String id) {
        //调用service层方法，查询市场活动
        Activity activity = activityService.queryActivityById(id);
        //根据查询结果，返回响应信息
        return activity;
    }

    @RequestMapping("/workbench/activity/saveEditActivity.do")
    @ResponseBody
    public Object saveEditActivity(Activity activity, HttpSession session) {
        User user = (User) session.getAttribute(Constants.SESSION_USER);
        //封装参数
        activity.setEditTime(DateUtils.formatDateTime(new Date()));
        activity.setEditBy(user.getId());

        ReturnObject returnObject = new ReturnObject();
        try {
            //调用service层方法，保存修改的市场活动
            int res = activityService.saveEditActivity(activity);
            if (res > 0) {
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_SUCCESS);
            } else {
                returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
                returnObject.setMessage("系统忙，请稍后重试");
            }

        } catch (Exception e) {
            e.printStackTrace();
            returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
            returnObject.setMessage("系统忙，请稍后重试");
        }

        return returnObject;
    }

    @RequestMapping("/workbench/activity/fileDownload.do")
    public void fileDownload(HttpServletResponse response) throws IOException {
        //1、设置响应类型
        response.setContentType("application/octet-stream;charset=UTF-8");
        //2、获取输出流
        ServletOutputStream out = response.getOutputStream();

        //浏览器接收到响应信息之后，默认情况下，直接在显示窗口中打开响应信息；
        //即使打不开，也会调用应用程序打开；只有实在打不开，才会激活文件下载窗口
        //可以设置响应头信息，使浏览器接收到响应信息之后，直接激活文件下载窗口，即使能打开也不打开
        response.addHeader("Content-Disposition", "attachment;filename=myStudentList.xls");

        //读取excel文件（InputStream），输出到浏览器（OutputStream）
        FileInputStream is = new FileInputStream("D:\\JAVAPRO\\crm-project\\temp\\studentList.xls");
        byte[] buffer = new byte[256];
        int len = 0;
        while ( (len = is.read(buffer)) != -1) {
            out.write(buffer, 0, len);
        }

        //关闭资源
        is.close();
        out.flush();
    }

    @RequestMapping("/workbench/activity/exportAllActivities.do")
    public void exportAllActivities(HttpServletResponse response) throws IOException {
        //调用service方法，查询所有市场活动
        List<Activity> activityList = activityService.queryAllActivities();
        //创建excel文件，并且把activityList写入到excel文件中
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("市场活动列表");
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("ID");
        cell = row.createCell(1);
        cell.setCellValue("所有者");
        cell = row.createCell(2);
        cell.setCellValue("名称");
        cell = row.createCell(3);
        cell.setCellValue("开始日期");
        cell = row.createCell(4);
        cell.setCellValue("结束日期");
        cell = row.createCell(5);
        cell.setCellValue("成本");
        cell = row.createCell(6);
        cell.setCellValue("描述");
        cell = row.createCell(7);
        cell.setCellValue("创建时间");
        cell = row.createCell(8);
        cell.setCellValue("创建者");
        cell = row.createCell(9);
        cell.setCellValue("修改时间");
        cell = row.createCell(10);
        cell.setCellValue("修改者");

        //遍历activityList，创建HSSFRow对象，生成所有的数据行
        if (activityList != null && activityList.size() > 0) {
            Activity activity = null;
            for(int i = 0; i < activityList.size(); i++) {
                activity = activityList.get(i);

                //每遍历出一个activity，生成一行
                row = sheet.createRow(i + 1);
                //每一行创建11列，每一列的数据从activity中获取
                cell = row.createCell(0);
                cell.setCellValue(activity.getId());
                cell = row.createCell(1);
                cell.setCellValue(activity.getOwner());
                cell = row.createCell(2);
                cell.setCellValue(activity.getName());
                cell = row.createCell(3);
                cell.setCellValue(activity.getStartDate());
                cell = row.createCell(4);
                cell.setCellValue(activity.getEndDate());
                cell = row.createCell(5);
                cell.setCellValue(activity.getCost());
                cell = row.createCell(6);
                cell.setCellValue(activity.getDescription());
                cell = row.createCell(7);
                cell.setCellValue(activity.getCreateTime());
                cell = row.createCell(8);
                cell.setCellValue(activity.getCreateBy());
                cell = row.createCell(9);
                cell.setCellValue(activity.getEditTime());
                cell = row.createCell(10);
                cell.setCellValue(activity.getEditBy());
            }
        }
        //根据wb对象生成excel文件
        /*OutputStream os = new FileOutputStream("D:\\JAVAPRO\\crm-project\\temp\\studentList.xls");
        wb.write(os);*/

        //关闭资源
        //os.close();
        //wb.close();

        //把生成的excel文件下载到客户端
        response.setContentType("application/octet-stream; charset=UTF-8");
        ServletOutputStream out = response.getOutputStream();
        response.addHeader("Content-Disposition", "attachment; filename=studentList.xls");
        /*FileInputStream is = new FileInputStream("D:\\JAVAPRO\\crm-project\\temp\\studentList.xls");
        byte[] buffer = new byte[256];
        int len = 0;
        while ((len = is.read(buffer)) != -1) {
            out.write(buffer, 0, len);
        }

        is.close();*/

        //直接从内存读取
        wb.write(out);
        wb.close();
        out.flush();
    }


    @RequestMapping("/workbench/activity/exportCheckedActivities.do")
    public void exportCheckedActivities(HttpServletResponse response, String[] id) throws IOException {
        //调用service方法，查询选择市场活动
        List<Activity> activityList = activityService.queryCheckedActivities(id);
        //创建excel文件，并且把activityList写入到excel文件中
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("市场活动列表");
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("ID");
        cell = row.createCell(1);
        cell.setCellValue("所有者");
        cell = row.createCell(2);
        cell.setCellValue("名称");
        cell = row.createCell(3);
        cell.setCellValue("开始日期");
        cell = row.createCell(4);
        cell.setCellValue("结束日期");
        cell = row.createCell(5);
        cell.setCellValue("成本");
        cell = row.createCell(6);
        cell.setCellValue("描述");
        cell = row.createCell(7);
        cell.setCellValue("创建时间");
        cell = row.createCell(8);
        cell.setCellValue("创建者");
        cell = row.createCell(9);
        cell.setCellValue("修改时间");
        cell = row.createCell(10);
        cell.setCellValue("修改者");

        //遍历activityList，创建HSSFRow对象，生成所有的数据行
        if (activityList != null && activityList.size() > 0) {
            Activity activity = null;
            for(int i = 0; i < activityList.size(); i++) {
                activity = activityList.get(i);

                //每遍历出一个activity，生成一行
                row = sheet.createRow(i + 1);
                //每一行创建11列，每一列的数据从activity中获取
                cell = row.createCell(0);
                cell.setCellValue(activity.getId());
                cell = row.createCell(1);
                cell.setCellValue(activity.getOwner());
                cell = row.createCell(2);
                cell.setCellValue(activity.getName());
                cell = row.createCell(3);
                cell.setCellValue(activity.getStartDate());
                cell = row.createCell(4);
                cell.setCellValue(activity.getEndDate());
                cell = row.createCell(5);
                cell.setCellValue(activity.getCost());
                cell = row.createCell(6);
                cell.setCellValue(activity.getDescription());
                cell = row.createCell(7);
                cell.setCellValue(activity.getCreateTime());
                cell = row.createCell(8);
                cell.setCellValue(activity.getCreateBy());
                cell = row.createCell(9);
                cell.setCellValue(activity.getEditTime());
                cell = row.createCell(10);
                cell.setCellValue(activity.getEditBy());
            }
        }
        //根据wb对象生成excel文件
        /*OutputStream os = new FileOutputStream("D:\\JAVAPRO\\crm-project\\temp\\studentList.xls");
        wb.write(os);*/

        //关闭资源
        //os.close();
        //wb.close();

        //把生成的excel文件下载到客户端
        response.setContentType("application/octet-stream; charset=UTF-8");
        ServletOutputStream out = response.getOutputStream();
        response.addHeader("Content-Disposition", "attachment; filename=studentList.xls");
        /*FileInputStream is = new FileInputStream("D:\\JAVAPRO\\crm-project\\temp\\studentList.xls");
        byte[] buffer = new byte[256];
        int len = 0;
        while ((len = is.read(buffer)) != -1) {
            out.write(buffer, 0, len);
        }

        is.close();*/

        //直接从内存读取
        wb.write(out);
        wb.close();
        out.flush();
    }

    /**
     * 配置springmvc文件上传解析器
     */
    @RequestMapping("/workbench/activity/fileUpload.do")
    @ResponseBody
    public Object fileUpload(String userName, MultipartFile myFile) throws IOException {
        //把文本数据打印到控制台
        System.out.println("userName=" + userName);
        //把文件在服务器指定的目录中生成一个同样的文件
        String originalFilename = myFile.getOriginalFilename();
        File file = new File("D:\\JAVAPRO\\crm-project\\temp\\", originalFilename);
        myFile.transferTo(file);

        //返回响应信息
        ReturnObject returnObject = new ReturnObject();
        returnObject.setCode(Constants.RETURN_OBJECT_CODE_SUCCESS);
        returnObject.setMessage("上传成功");
        return returnObject;
    }

    @RequestMapping("/workbench/activity/importActivity.do")
    @ResponseBody
    public Object importActivity(MultipartFile activityFile, HttpSession session) {

        User user = (User) session.getAttribute(Constants.SESSION_USER);
        ReturnObject returnObject = new ReturnObject();
        try {
            //把excel文件写入磁盘目录
            //String originalFilename = activityFile.getOriginalFilename();
            //File file = new File("D:\\JAVAPRO\\crm-project\\temp\\", originalFilename);
            //activityFile.transferTo(file);

            //解析excel文件，获取文件中的数据，并且封装成activityList
            //InputStream is = new FileInputStream("D:\\JAVAPRO\\crm-project\\temp\\" +  originalFilename);

            InputStream is = activityFile.getInputStream();

            HSSFWorkbook wb = new HSSFWorkbook(is);
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row = null;
            HSSFCell cell = null;
            Activity activity = null;
            List<Activity> activityList = new ArrayList<>();
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                row = sheet.getRow(i);
                activity = new Activity();
                activity.setId(UUIDUtils.getUUID());
                activity.setOwner(user.getId());
                activity.setCreateTime(DateUtils.formatDateTime(new Date()));
                activity.setCreateBy(user.getId());

                for (int j = 0; j < row.getLastCellNum(); j++) {
                    cell = row.getCell(j);
                    String cellValue = HSSFUtils.getCellValueForStr(cell);
                    if (j == 0) {
                        activity.setName(cellValue);
                    } else if (j == 1) {
                        activity.setStartDate(cellValue);
                    } else if (j == 2) {
                        activity.setEndDate(cellValue);
                    } else if (j == 3) {
                        activity.setCost(cellValue);
                    } else if (j == 4) {
                        activity.setDescription(cellValue);
                    }
                }

                //每一行中所有列都封装完成之后，把activity保存在list中
                activityList.add(activity);
            }

            //调用service层方法，保存市场活动
            int ret = activityService.saveCreateActivityByList(activityList);
            returnObject.setCode(Constants.RETURN_OBJECT_CODE_SUCCESS);
            returnObject.setRetData(ret);

        } catch (IOException e) {
            e.printStackTrace();
            returnObject.setCode(Constants.RETURN_OBJECT_CODE_FAIL);
            returnObject.setMessage("系统忙，请稍后重试...");
        }
        return returnObject;
    }

    @RequestMapping("/workbench/activity/detailActivity.do")
    public String detailActivity(String id, HttpServletRequest request){
        //调用service方法
        Activity activity = activityService.queryActivityForDetailById(id);
        List<ActivityRemark> remarkList = activityRemarkService.queryActivityRemarkForDetailByActivityId(id);
        request.setAttribute("activity", activity);
        request.setAttribute("remarkList", remarkList);
        //请求转发
        return "workbench/activity/detail";
    }


}
