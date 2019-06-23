package com.example.poi.service;

import com.example.poi.mapper.UserMapper;
import com.example.poi.model.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class UserService {

    @Autowired
    private UserMapper userMapper;

    public int addUser(User user){
        return userMapper.addUser(user);
    }

    public User findById(int id){
        return userMapper.findById(id);
    }

    public List<User> listUser(){
        return userMapper.listUser();
    }

    public static void main(String[] args) {
        Cell cell=null;
        Sheet sheet=null;
    }
}
