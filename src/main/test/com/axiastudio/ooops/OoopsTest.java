package com.axiastudio.ooops;

import com.axiastudio.ooops.filters.Filter;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class OoopsTest {

    @Test
    public void test() throws Exception {

        List<List<String>> books = new ArrayList<>();
        books.add(null); // skip first line
        books.add(new ArrayList<>(Arrays.asList("Anna Karenina", "Tolstoy")));
        books.add(new ArrayList<>(Arrays.asList("The Master and Margarita", "Bulgakov")));

        Ooops.create()
                .open("uno:socket,host=localhost,port=8997;urp;StarOffice.ServiceManager")
                .load(new FileInputStream(new File("test.odt")))
                .fillBookmark("name", "Tiziano")
                .fillTable("bookTable", books)
                .showHideSection("appear", true)
                .showHideSection("notAppear", false)
                .filter(Filter.PDF)
                .toStream(new FileOutputStream(new File("test.pdf")));

    }

}