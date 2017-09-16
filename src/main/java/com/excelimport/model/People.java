package com.excelimport.model;


import lombok.Getter;
import lombok.Setter;
import lombok.experimental.Accessors;

/**me
 * Created by oguz on 16.09.2017.
 */
@Accessors(chain = true)
public class People {

    @Getter
    @Setter
    private String name;

    @Getter
    @Setter
    private String lastName;

}
