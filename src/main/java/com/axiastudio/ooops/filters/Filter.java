package com.axiastudio.ooops.filters;

/**
 * User: tiziano
 * Date: 03/06/14
 * Time: 16:59
 */
public enum Filter {

    ODT("writer8"), PDF("writer_pdf_Export"), DOC("MS Word 97");

    private String filterName;

    Filter(String filterName) {
        this.filterName = filterName;
    }

    public String filterName() {
        return filterName;
    }

}
