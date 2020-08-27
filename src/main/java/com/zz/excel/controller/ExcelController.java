package com.zz.excel.controller;

import com.zz.excel.entity.ExcelEntity;
import com.zz.excel.service.ExcelService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.List;

@Controller
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    private Logger logger = LoggerFactory.getLogger(this.getClass());

    @RequestMapping(value = "/upload", method = RequestMethod.GET)
    @ResponseBody
    public String upload(@RequestParam("file")MultipartFile file){
        List<ExcelEntity> list = excelService.readExcel(file);
        System.out.println(list.toString());
        return "success";
    }

    @RequestMapping(value = "/upload1", method = RequestMethod.GET)
    @ResponseBody
    public ModelAndView upload1(Model model){
        return new ModelAndView("upload");
    }

    @RequestMapping(value = "/write", method = RequestMethod.GET)
    @ResponseBody
    public String write(){
        ExcelEntity entity = new ExcelEntity();
        entity.setName("张三");
        entity.setAge(12);
        entity.setPhone("1456852");
        entity.setEmail("123");
        excelService.writeExcel(entity);
        return "success";
    }

    @RequestMapping(value = "/download", method = RequestMethod.GET, produces = MediaType.ALL_VALUE)
    public void download(HttpServletResponse response) throws IOException {
        ExcelEntity entity = new ExcelEntity();
        entity.setName("张三");
        entity.setAge(12);
        entity.setPhone("1456852");
        entity.setEmail("123");
        HSSFWorkbook workbook = excelService.downloadExcel(entity);
        response.setHeader("Content-Disposition", "attachment; filename=information.xls");
        workbook.write(response.getOutputStream());
    }


    @RequestMapping(value = "/testzz", method = RequestMethod.GET)
    public void getIp(HttpServletRequest request){
        logger.info("remote addr:" + request.getRemoteAddr());
        logger.info("X-Forwarded-For :" + request.getHeader("X-Forwarded-For"));
        logger.info("Proxy-Client-IP :" + request.getHeader("Proxy-Client-IP"));
        logger.info("WL-Proxy-Client-IP :" + request.getHeader("WL-Proxy-Client-IP"));
        logger.info("HTTP_X_FORWARDED_FOR :" + request.getHeader("HTTP_X_FORWARDED_FOR"));
        logger.info("HTTP_CLIENT_IP :" + request.getHeader("HTTP_CLIENT_IP"));
    }
}
