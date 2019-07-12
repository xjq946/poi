package com.example.poi.controller;

import com.example.poi.model.User;
import com.example.poi.service.UserService;
import com.example.poi.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.util.CollectionUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.List;
import java.util.Objects;
import java.util.UUID;

@RestController
@RequestMapping("/user")
public class UserController {

    @Autowired
    private UserService userService;

    @RequestMapping("/addUser")
    public String addUser(@RequestBody User user){
        userService.addUser(user);
        return "insert success!";
    }

    @RequestMapping("/findById/{id}")
    public User findById(@PathVariable("id") Integer id){
        User user = userService.findById(id);
        return user;
    }

    @RequestMapping("/listUser")
    public List<User> listUser(){
        List<User> userList = userService.listUser();
        return userList;
    }

    /**
     * 导入文件
     * @param multipartFile
     * @return
     */
    @RequestMapping("/importExcel")
    public String importExcel(@RequestParam("file") MultipartFile multipartFile){
        BufferedOutputStream bos=null;
        try {
            //保存文件
            String originalFilename = multipartFile.getOriginalFilename();
            String filename= UUID.randomUUID().toString().replace("-","").toLowerCase()+"_"+originalFilename;
            bos=new BufferedOutputStream(new FileOutputStream(filename));
            Workbook workbook= WorkbookFactory.create(multipartFile.getInputStream());
            workbook.write(bos);

            //读取保存后的文件的内容
            File file=new File(filename);
            String absolutePath = file.getAbsolutePath();
            List<User> userList = ExcelUtils.readExcel(absolutePath);
            if(Objects.nonNull(userList)){
                for(User user:userList){
                    userService.addUser(user);
                }
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            if(bos!=null){
                try {
                    bos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return "导入成功！";
    }

    /**
     * 下载文件
     * @param filename
     * @return
     */
    @RequestMapping("/export/{filename}")
    public String export(@PathVariable("filename") String filename, HttpServletResponse response){
        try {
            //加载模板
            InputStream inp=new FileInputStream("myTemplate/"+filename);
            Workbook workbook=WorkbookFactory.create(inp);
            //查询数据
            List<User> userList = userService.listUser();

            Sheet sheet = workbook.getSheetAt(0);

            //填充数据
            if(Objects.nonNull(userList)&& !CollectionUtils.isEmpty(userList)){
              for(int i=0;i<userList.size();i++){
                  Row row = sheet.createRow(i+1);
                  ExcelUtils.setCellData(userList.get(i),row,sheet);
              }
            }
            //下载
            ExcelUtils.downLoadExcel(filename,response,workbook);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return "导出成功！";
    }

}
