package com.example.excelconverterplugin.domain;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.time.Instant;

@Setter
@Getter
@ToString
public class UserData {
    private String fullName;
    private Integer age;
    private Instant birthday;
}
