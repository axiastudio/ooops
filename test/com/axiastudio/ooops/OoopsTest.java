package com.axiastudio.ooops;

import com.axiastudio.ooops.filters.Filter;
import org.junit.Test;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class OoopsTest {

    @Test
    public void test() throws Exception {

        Map<String, Object> map = new HashMap<>();
        map.put("name", "Tiziano");

        Ooops.create()
                .open("uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager")
                .load(new FileInputStream(new File("test.odt")))
                .map(map)
                .filter(Filter.writer_pdf_Export)
                .toStream(new FileOutputStream(new File("test.pdf")));

    }

}