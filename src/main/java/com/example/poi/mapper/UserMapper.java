package com.example.poi.mapper;

import com.example.poi.model.User;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import java.util.List;
@Mapper
public interface UserMapper {

    int addUser(User user);

    User findById(int id);

    List<User> listUser();
}
