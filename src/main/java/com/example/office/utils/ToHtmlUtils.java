package com.example.office.utils;

import com.example.office.exception.RuntimeCodeException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.convert.out.ConversionFeatures;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.Map;

public class ToHtmlUtils {


    final  static boolean save;
    final  static boolean nestLists;
    static {
        save = true;
        nestLists = true;
    }

    public static void convertDocxToHtml(String  docxFilePath , String htmlPath) {
        try {
            long start = System.currentTimeMillis();
            WordprocessingMLPackage wordMLPackage = Docx4J.load(new java.io.File(docxFilePath));

            HTMLSettings htmlSettings = Docx4J.createHTMLSettings();

            htmlSettings.setImageDirPath(htmlPath + "_images");
           // htmlSettings.setImageTargetUri(docxFilePath+"_images");
            htmlSettings.setImageTargetUri(htmlPath.substring(htmlPath.lastIndexOf("/")+1) + "_images");
            htmlSettings.setWmlPackage(wordMLPackage);
            String  userCSS = "tohtml, body, div, span, h1, h2, h3, h4, h5, h6, p, a, img,  ol, ul, li, table, caption, tbody, tfoot, thead, tr, th, td " +
                    "{ margin: 0; padding: 0; border: 0;}" +
                    "body {line-height: 1;} ";

            htmlSettings.setUserCSS(userCSS);
            htmlSettings.getFeatures().remove(ConversionFeatures.PP_HTML_COLLECT_LISTS);
            OutputStream os;
            if (save) {
                os = new FileOutputStream(htmlPath+".html");
            } else {
                os = new ByteArrayOutputStream();
            }
            Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.OutputMethodXML", true);
            Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
            if (save) {
                System.out.println("Saved: " + htmlPath + ".html ");
            } else {
                System.out.println( ((ByteArrayOutputStream)os).toString() );
            }
            // Clean up, so any ObfuscatedFontPart temp files can be deleted
            if (wordMLPackage.getMainDocumentPart().getFontTablePart()!=null) {
                wordMLPackage.getMainDocumentPart().getFontTablePart().deleteEmbeddedFontTempFiles();
            }
            os.close();
        }catch (Exception e){
            throw new RuntimeCodeException("系统错误");
        }
    }


    public static void PdfToImage(String pdfurl ,String htmlPath){
        StringBuffer buffer = new StringBuffer();
        FileOutputStream fos = null;
        PDDocument document = null;
        File pdfFile;
        int size;
        BufferedImage image;
        FileOutputStream out = null;
        Long randStr = 0l;
        //PDF转换成HTML保存的文件夹
        //本机当做服务器
        File htmlsDir = new File(htmlPath);
        if(!htmlsDir.exists()){
            htmlsDir.mkdirs();
        }
        File htmlDir = new File(htmlPath+"/");
        if(!htmlDir.exists()){
            htmlDir.mkdirs();
        }
        try{
            //遍历处理pdf附件
            randStr = System.currentTimeMillis();
            buffer.append("<!doctype html>\r\n");
            buffer.append("<head>\r\n");
            buffer.append("<meta charset=\"UTF-8\">\r\n");
            buffer.append("</head>\r\n");
            buffer.append("<body style=\"background-color:gray;\">\r\n");
            buffer.append("<style>\r\n");
            buffer.append("img {background-color:#fff; text-align:center; width:100%; max-width:100%;margin-top:6px;}\r\n");
            buffer.append("</style>\r\n");
            document = new PDDocument();
            //pdf附件
            pdfFile = new File(pdfurl);
            document = PDDocument.load(pdfFile, (String) null);
            /*加水印代码
            document.setAllSecurityToBeRemoved(true);
            for(PDPage page:document.getPages()){
                PDPageContentStream cs = new PDPageContentStream(document, page, true, true, true);
                String ts = "Some sample text";
                PDFont font = PDType1Font.HELVETICA_OBLIQUE;
                float fontSize = 50.0f;
                PDResources resources = page.getResources();
                PDExtendedGraphicsState r0 = new PDExtendedGraphicsState();
                // 透明度
                r0.setNonStrokingAlphaConstant(0.2f);
                r0.setAlphaSourceFlag(true);
                cs.setGraphicsStateParameters(r0);
                cs.setNonStrokingColor(200,0,0);//Red
                cs.beginText();
                cs.setFont(font, fontSize);
                // 获取旋转实例
                cs.setTextMatrix(Matrix.getRotateInstance(20,350f,490f));
                cs.showText(ts);
                cs.endText();

                cs.close();
            }
            document.save(pdfFile);*/

            size = document.getNumberOfPages();
            Long start = System.currentTimeMillis(), end = null;
            System.out.println("===>pdf : " + pdfFile.getName() +" , size : " + size);
            PDFRenderer reader = new PDFRenderer(document);
            for(int i=0 ; i < size; i++){
                //image = new PDFRenderer(document).renderImageWithDPI(i,130,ImageType.RGB);
                image = reader.renderImage(i, 1.5f);
                //生成图片,保存位置
                out = new FileOutputStream(htmlsDir + "/"+ "image" + "_" + i + ".png");

                ImageIO.write(image, "png", out); //使用png的清晰度
                //将图片路径追加到网页文件里
                buffer.append("<img src=\"" + htmlsDir +"/"+ "image" + "_" + i + ".png\"/>\r\n");
                image = null; out.flush(); out.close();
            }
            reader = null;
            document.close();
            buffer.append("</body>\r\n");
            buffer.append("</html>");
            end = System.currentTimeMillis() - start;
            System.out.println("===> Reading pdf times: " + (end/1000));
            start = end = null;
            //生成网页文件
            fos = new FileOutputStream(htmlsDir+".html");
            System.out.println(htmlsDir+".html");
            fos.write(buffer.toString().getBytes());
            fos.flush(); fos.close();
            buffer.setLength(0);



        }catch(Exception e){
            System.out.println("===>Reader parse pdf to jpg error : " + e.getMessage());
            e.printStackTrace();
        }finally {
            if(out !=null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(document != null){
                try {
                    document.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(fos != null){
                try {
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * excel to html
     */
    public static void excelToHtml(String filePath,String htmlFile, Map<String, String> infoMap) {
        FileWriter writer = null;
        try {
            String result = POIReadExcel.readExcelToHtml(new FileInputStream(filePath), infoMap);
            writer = new FileWriter(htmlFile);
            writer.write(result);
            writer.write("<script>;(function (win, lib) {\n" +
                    "  var doc = win.document\n" +
                    "  var docEl = doc.documentElement\n" +
                    "  var metaEl = doc.querySelector('meta[name=\"viewport\"]')\n" +
                    "  var flexibleEl = doc.querySelector('meta[name=\"flexible\"]')\n" +
                    "  var dpr = 0\n" +
                    "  var scale = 0\n" +
                    "  var tid\n" +
                    "  var flexible = lib.flexible || (lib.flexible = {})\n" +
                    "\n" +
                    "  if (metaEl) {\n" +
                    "    console.warn('将根据已有的meta标签来设置缩放比例')\n" +
                    "    var match = metaEl.getAttribute('content').match(/initial-scale=([\\d.]+)/)\n" +
                    "    if (match) {\n" +
                    "      scale = parseFloat(match[1])\n" +
                    "      dpr = parseInt(1 / scale)\n" +
                    "    }\n" +
                    "  } else if (flexibleEl) {\n" +
                    "    var content = flexibleEl.getAttribute('content')\n" +
                    "    if (content) {\n" +
                    "      var initialDpr = content.match(/initial-dpr=([\\d.]+)/)\n" +
                    "      var maximumDpr = content.match(/maximum-dpr=([\\d.]+)/)\n" +
                    "      if (initialDpr) {\n" +
                    "        dpr = parseFloat(initialDpr[1])\n" +
                    "        scale = parseFloat((1 / dpr).toFixed(2))\n" +
                    "      }\n" +
                    "      if (maximumDpr) {\n" +
                    "        dpr = parseFloat(maximumDpr[1])\n" +
                    "        scale = parseFloat((1 / dpr).toFixed(2))\n" +
                    "      }\n" +
                    "    }\n" +
                    "  }\n" +
                    "\n" +
                    "  if (!dpr && !scale) {\n" +
                    "    var isIPhone = win.navigator.appVersion.match(/iphone/gi)\n" +
                    "    var devicePixelRatio = win.devicePixelRatio\n" +
                    "    if (isIPhone) {\n" +
                    "      // iOS下，对于2和3的屏，用2倍的方案，其余的用1倍方案\n" +
                    "      if (devicePixelRatio >= 3 && (!dpr || dpr >= 3)) {\n" +
                    "        dpr = 3\n" +
                    "      } else if (devicePixelRatio >= 2 && (!dpr || dpr >= 2)) {\n" +
                    "        dpr = 2\n" +
                    "      } else {\n" +
                    "        dpr = 1\n" +
                    "      }\n" +
                    "    } else {\n" +
                    "      // 其他设备下，仍旧使用1倍的方案\n" +
                    "      dpr = 1\n" +
                    "    }\n" +
                    "    scale = 1 / dpr\n" +
                    "  }\n" +
                    "\n" +
                    "  docEl.setAttribute('data-dpr', dpr)\n" +
                    "  if (!metaEl) {\n" +
                    "    metaEl = doc.createElement('meta')\n" +
                    "    metaEl.setAttribute('name', 'viewport')\n" +
                    "    metaEl.setAttribute('content', 'initial-scale=' + scale + ', maximum-scale=' + scale + ', minimum-scale=' + scale + ', user-scalable=no')\n" +
                    "    if (docEl.firstElementChild) {\n" +
                    "      docEl.firstElementChild.appendChild(metaEl)\n" +
                    "    } else {\n" +
                    "      var wrap = doc.createElement('div')\n" +
                    "      wrap.appendChild(metaEl)\n" +
                    "      doc.write(wrap.innerHTML)\n" +
                    "    }\n" +
                    "  }\n" +
                    "\n" +
                    "  function refreshRem () {\n" +
                    "    var width = docEl.getBoundingClientRect().width\n" +
                    "    if (width / dpr > 540) {\n" +
                    "      width = 540 * dpr\n" +
                    "    }\n" +
                    "    var rem = width / 10\n" +
                    "    docEl.style.fontSize = rem + 'px'\n" +
                    "    flexible.rem = win.rem = rem\n" +
                    "  }\n" +
                    "\n" +
                    "  win.addEventListener('resize', function () {\n" +
                    "    clearTimeout(tid)\n" +
                    "    tid = setTimeout(refreshRem, 300)\n" +
                    "  }, false)\n" +
                    "  win.addEventListener('pageshow', function (e) {\n" +
                    "    if (e.persisted) {\n" +
                    "      clearTimeout(tid)\n" +
                    "      tid = setTimeout(refreshRem, 300)\n" +
                    "    }\n" +
                    "  }, false)\n" +
                    "\n" +
                    "  if (doc.readyState === 'complete') {\n" +
                    "    doc.body.style.fontSize = 12 * dpr + 'px'\n" +
                    "  } else {\n" +
                    "    doc.addEventListener('DOMContentLoaded', function (e) {\n" +
                    "      doc.body.style.fontSize = 12 * dpr + 'px'\n" +
                    "    }, false)\n" +
                    "  }\n" +
                    "\n" +
                    "  refreshRem()\n" +
                    "\n" +
                    "  flexible.dpr = win.dpr = dpr\n" +
                    "  flexible.refreshRem = refreshRem\n" +
                    "  flexible.rem2px = function (d) {\n" +
                    "    var val = parseFloat(d) * this.rem\n" +
                    "    if (typeof d === 'string' && d.match(/rem$/)) {\n" +
                    "      val += 'px'\n" +
                    "    }\n" +
                    "    return val\n" +
                    "  }\n" +
                    "  flexible.px2rem = function (d) {\n" +
                    "    var val = parseFloat(d) / this.rem\n" +
                    "    if (typeof d === 'string' && d.match(/px$/)) {\n" +
                    "      val += 'rem'\n" +
                    "    }\n" +
                    "    return val\n" +
                    "  }\n" +
                    "})(window, window['lib'] || (window['lib'] = {}))</script>");
            writer.write("<script>let title = document.getElementById('sheet');\n" +
                    "let content = document.getElementById('content');\n" +
                    "title.onclick = event =>{\n" +
                    "  let e = event || window.event;\n" +
                    "  let target = e.target || e.srcElement;\n" +
                    "  if (target.nodeName.toLowerCase() == 'span') {\n" +
                    "    for(var i = 0; i < title.children.length; i++){\n" +
                    "        title.children[i].className = '';\n" +
                    "        content.children[i].style.display = 'none';\n" +
                    "        if(title.children[i] === target){\n" +
                    "            content.children[i].style.display = 'block';\n" +
                    "            var thead = content.children[i].children[0].children[0];\n" +
                    "            if(thead.tagName.toLowerCase() === 'thead'){\n" +
                    "                thead.style.transform = 'translate(0, '+ (content.scrollTop-1)+'px)';\n" +
                    "            }\n" +
                    "        }\n" +
                    "    }\n" +
                    "    target.className = 'active';\n" +
                    "  }\n" +
                    "}\n" +
                    "\n" +
                    "content.addEventListener('scroll', function () {\n" +
                    "  var top = this.scrollTop;\n" +
                    "  for(var j = 0; j < title.children.length; j++){\n" +
                    "    if (title.children[j].className === 'active') {\n" +
                    "      var thead = content.children[j].children[0].children[0];\n" +
                    "      if (thead.tagName.toLowerCase() === 'thead') {\n" +
                    "          thead.style.transform = 'translate(0, '+ (top-1)+'px)';\n" +
                    "      }\n" +
                    "    }\n" +
                    "  }\n" +
                    "})\n" +
                    "\n" +
                    "\n" +
                    "\n" +
                    "\n" +
                    "\n" +
                    "\n</script>");
            writer.write("</body>");
            writer.flush();
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            try {
                if(writer!=null)
                    writer.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static String GetFileExt(String path) {
        String ext = null;
        int i = path.lastIndexOf('.');
        if (i > 0 && i < path.length() - 1) {
            ext = path.substring(i + 1).toLowerCase();
        }
        return ext;
    }
}
