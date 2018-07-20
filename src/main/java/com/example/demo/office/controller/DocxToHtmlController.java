package com.example.demo.office.controller;

import com.example.demo.office.utils.DocxToHtmlUtils;
import com.example.demo.office.utils.ReturnMassage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

@Controller
public class DocxToHtmlController {

    private Logger logger = LoggerFactory.getLogger(this.getClass().getName());


    @RequestMapping("file")
    public String file(){
        return "/file";
    }


    @RequestMapping(value="/officeToHtml", method = RequestMethod.POST)
    @ResponseBody
    public ReturnMassage officeToHtml( HttpServletResponse response,
                                                    HttpServletRequest request,
                                                    @RequestParam("file") MultipartFile  file)
            {
            logger.info(file.getOriginalFilename());

        try {
            String path = file.getOriginalFilename();
            if(null == path){
                return ReturnMassage.ok().put("ok","文件路径错误！");
            }
            String hPath = path.substring(path.lastIndexOf("\\")+1);
            String ext = DocxToHtmlUtils.GetFileExt(path);
            String htmlPath = System.getProperty("user.dir") +"\\html\\"+hPath;

            if("docx".equals(ext)){
                DocxToHtmlUtils.convertDocxToHtml(path,htmlPath);
            }else if("pdf".equals(ext)){
                DocxToHtmlUtils.PdfToImage(path,htmlPath);
            }else if("xlsx".equals(ext)){
                //String wPath = "E:\\docx4j-test\\docxToHtml";
                Map<String, String> infoMap = new HashMap<String, String>();
                // 上传HTML 服务器路径， 暂时为本地
                infoMap.put("uploadFile",htmlPath);
                infoMap.put("readfile", htmlPath);
                File demoFile = new File(htmlPath,"demo.html");
                if(!demoFile.exists()) {
                    demoFile.getParentFile().mkdirs();
                    demoFile.createNewFile();
                }
                DocxToHtmlUtils.excelToHtml(path,demoFile.getAbsolutePath(),infoMap);
                htmlPath = htmlPath + "\\demo";
            }else{
                return ReturnMassage.ok().put("ok","不支持该类型文件！");
            }

            File openFile = new File(htmlPath+".html");
            Runtime ce=Runtime.getRuntime();
            System.out.println(openFile.getAbsolutePath());

            ce.exec("rundll32 url.dll,FileProtocolHandler " + openFile.getAbsolutePath());

        } catch (Exception e) {
                //Exception 等待处理
                e.printStackTrace();
        }finally {

        }
        return  ReturnMassage.ok();
    }
}
