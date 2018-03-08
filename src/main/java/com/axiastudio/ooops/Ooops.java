package com.axiastudio.ooops;

import com.axiastudio.ooops.filters.Filter;
import com.axiastudio.ooops.streams.OoopsInputStream;
import com.axiastudio.ooops.streams.OoopsOutputStream;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.frame.*;
import com.sun.star.io.IOException;
import com.sun.star.lang.*;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.text.*;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;
import com.sun.star.view.XSelectionSupplier;

import java.io.*;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * User: tiziano
 * Date: 03/06/14
 * Time: 16:15
 */
public class Ooops {

    private XComponentLoader loader = null;
    private XComponent component = null;
    private XDispatchHelper dispatchHelper = null;
    private Filter filter = Filter.ODT;

    public static Ooops create() {
        return new Ooops();
    }

    public Ooops open(String connectionString){
        try{
            XComponentContext context = com.sun.star.comp.helper.Bootstrap.createInitialComponentContext(null);
            XMultiComponentFactory serviceManager = context.getServiceManager();
            Object objResolver = serviceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", context);
            XUnoUrlResolver resolver = UnoRuntime.queryInterface(XUnoUrlResolver.class, objResolver);
            Object objectInitial = resolver.resolve(connectionString);
            XMultiComponentFactory factory = UnoRuntime.queryInterface(XMultiComponentFactory.class, objectInitial);
            XPropertySet properties = UnoRuntime.queryInterface(XPropertySet.class, factory);
            Object objContext = properties.getPropertyValue("DefaultContext");
            context = UnoRuntime.queryInterface(XComponentContext.class, objContext);
            loader = UnoRuntime.queryInterface(XComponentLoader.class, factory.createInstanceWithContext("com.sun.star.frame.Desktop", context));
            dispatchHelper = UnoRuntime.queryInterface(XDispatchHelper.class, factory.createInstanceWithContext("com.sun.star.frame.DispatchHelper", context));
        } catch (java.lang.Exception ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, "Unable to open connection to the listener with connection string '" + connectionString + "'.");
            return null;
        }
        return this;
    }

    private void blank(){
        try {
            PropertyValue[] propertyvalue = new PropertyValue[0];
            component = loader.loadComponentFromURL("private:factory/swriter", "_blank", 0, propertyvalue);
        } catch (IOException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, null, ex);
        } catch (com.sun.star.lang.IllegalArgumentException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public Ooops load(String url){
        try {
            component = loader.loadComponentFromURL(url, "_blank", 0, new PropertyValue[0]);
        } catch (com.sun.star.io.IOException e) {
            e.printStackTrace();
        } catch (IllegalArgumentException e) {
            e.printStackTrace();
        }
        return this;
    }


    public Ooops load(InputStream inputStream){
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        try {
            byte[] byteBuffer = new byte[4096];
            int byteBufferLength;
            while ((byteBufferLength = inputStream.read(byteBuffer)) > 0) {
                bytes.write(byteBuffer, 0, byteBufferLength);
            }
            inputStream.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, null, ex);
        } catch (java.io.IOException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, null, ex);
        }
        byte[] buf = bytes.toByteArray();
        OoopsInputStream ooopsInputStream = new OoopsInputStream(buf);
        component = loadComponent(ooopsInputStream);
        return this;
    }

    public Ooops load(byte[] content){
        /*
        OfficeInputStream officeInputStream = new OfficeInputStream(content);
        component = loadComponent(officeInputStream);
        */
        InputStream inputStream = new ByteArrayInputStream(content);
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        try {
            byte[] byteBuffer = new byte[4096];
            int byteBufferLength;
            while ((byteBufferLength = inputStream.read(byteBuffer)) > 0) {
                bytes.write(byteBuffer, 0, byteBufferLength);
            }
            inputStream.close();
        } catch (java.io.IOException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, null, ex);
        }
        byte[] buf = bytes.toByteArray();
        OoopsInputStream ooopsInputStream = new OoopsInputStream(buf);
        component = loadComponent(ooopsInputStream);
        return this;
    }

    private XComponent loadComponent(OoopsInputStream inStream){
        Boolean hidden = Boolean.TRUE;
        try {
            PropertyValue[] propertyValue = new PropertyValue[2];
            propertyValue[0] = new PropertyValue();
            propertyValue[0].Name = "InputStream";
            propertyValue[0].Value = inStream;
            if( hidden ){
                propertyValue[1] = new PropertyValue();
                propertyValue[1].Name = "Hidden";
                propertyValue[1].Value = new Boolean(true);
            } else {
                propertyValue[1] = new PropertyValue();
                propertyValue[1].Name = "Hidden";
                propertyValue[1].Value = new Boolean(false);
            }
            try {
                return loader.loadComponentFromURL("private:stream", "_blank", 0, propertyValue);
            } catch (IllegalArgumentException e) {
                e.printStackTrace();
            }
        } catch (com.sun.star.io.IOException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, null, ex);
        }
        return null;
    }


    /*
     *   fills
     */
    public Ooops fillBookmark(String name, String value){
        XTextRange anchor = getTextRange(name, component);
        if( anchor != null &&  value != null ){
            anchor.setString(value.toString());
            //setBookmark(name, anchor); // reset the original bookmark
        }
        return this;
    }

    public Ooops showHideSection(String name, Boolean show){
        XTextSection section = getSection(name, component);
        if( section != null ){
            XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, section);
            try {
                xPropertySet.setPropertyValue("IsVisible", show);
            } catch (UnknownPropertyException e) {
                e.printStackTrace();
            } catch (PropertyVetoException e) {
                e.printStackTrace();
            } catch (IllegalArgumentException e) {
                e.printStackTrace();
            } catch (WrappedTargetException e) {
                e.printStackTrace();
            }
        }
        return this;
    }

    private void setBookmark(String name, XTextRange anchor) {
        XText text = anchor.getText();
        XTextCursor cursor = text.createTextCursorByRange(anchor);
        XTextDocument document = UnoRuntime.queryInterface(XTextDocument.class, component);
        XMultiServiceFactory factory = UnoRuntime.queryInterface(XMultiServiceFactory.class, document);
        Object bookmark;
        try {
            bookmark = factory.createInstance("com.sun.star.text.Bookmark");
            XNamed named = UnoRuntime.queryInterface(XNamed.class, bookmark);
            named.setName(name);
            XText documentText = cursor.getText();
            XTextContent textContent = UnoRuntime.queryInterface(XTextContent.class, bookmark);
            documentText.insertTextContent(cursor, textContent, Boolean.TRUE);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String columnName(Integer idx) {
        StringBuilder s = new StringBuilder();
        while( idx>=26 ){
            s.insert(0, (char) ('A' + idx % 26));
            idx = idx / 26 - 1;
        }
        s.insert(0, (char) ('A' + idx));
        return s.toString();
    }

    public Ooops removeTable(String name){
        XTextTable textTable = getTextTable(name, component);
        XModel model = UnoRuntime.queryInterface(XModel.class, component);

        XTextDocument textDocument = UnoRuntime.queryInterface(XTextDocument.class, model);
        XText text = textDocument.getText();
        try {
            text.removeTextContent(UnoRuntime.queryInterface(XTextContent.class, textTable));
        } catch (NoSuchElementException e) {
            e.printStackTrace();
        }
        return this;
    }

    public Ooops fillTable(String name, List<List<String>> values){
        if( values==null ){
            removeTable(name);
        }
        XTextTable textTable = getTextTable(name, component);

        XModel model = UnoRuntime.queryInterface(XModel.class, component);
        XController controller = model.getCurrentController();
        XSelectionSupplier supplier = UnoRuntime.queryInterface(XSelectionSupplier.class, controller);

        // append columns
        Integer numOfRowsInTable = textTable.getRows().getCount();
        Integer numOfCols;
        if( values.get(0)==null ){
            // skip headers
            numOfCols = values.get(1).size();
        }else {
            numOfCols = values.get(0).size();
        }
        int initNumOfCols = textTable.getColumns().getCount();
        for( Integer i=0; i<(numOfCols-initNumOfCols); i++ ){
            XCellRange textTableCellRange = UnoRuntime.queryInterface(XCellRange.class, textTable);
            String copyRangeName = columnName(i) + "1:" + columnName(i) + numOfRowsInTable.toString();
            XCellRange copyCellRange = textTableCellRange.getCellRangeByName(copyRangeName);
            try {
                supplier.select(copyCellRange);
                dispatchHelper.executeDispatch(UnoRuntime.queryInterface(XDispatchProvider.class, controller.getFrame()), ".uno:Copy", "", 0, new PropertyValue[]{new PropertyValue()});
            } catch (IllegalArgumentException e) {
                e.printStackTrace();
            }
            textTable.getColumns().insertByIndex(textTable.getColumns().getCount(), 1);
            String lastColumnName = columnName(textTable.getColumns().getCount()-1);
            String lastColumnRangeName = lastColumnName + "1:" + lastColumnName + numOfRowsInTable.toString();
            XCellRange pasteCellRange = textTableCellRange.getCellRangeByName(lastColumnRangeName);
            try {
                supplier.select(pasteCellRange);
                dispatchHelper.executeDispatch(UnoRuntime.queryInterface(XDispatchProvider.class, controller.getFrame()), ".uno:Paste", "", 0, null);
            } catch (IllegalArgumentException e) {
                e.printStackTrace();
            }
        }

        // append rows
        int numOfRows = values.size();
        if( numOfRows>numOfRowsInTable ) {
            textTable.getRows().insertByIndex(numOfRowsInTable, numOfRows - numOfRowsInTable);
        }
        Integer r=1;
        for( List<String> row: values ){
            Integer c=0;
            if( row!=null ) {
                for (String value : row) {
                    XCell cell = textTable.getCellByName(columnName(c) + r.toString());
                    XText text = UnoRuntime.queryInterface(XText.class, cell);
                    if( text!=null ) {
                        text.setString(value);
                    }
                    c++;
                }
            }
            r++;
        }
        return this;
    }



    /*
     *   discovers
     */
    private XTextRange getTextRange(String bookmarkName, XComponent component){
        XBookmarksSupplier supplier = UnoRuntime.queryInterface(XBookmarksSupplier.class, component);
        XNameAccess bookmarks = supplier.getBookmarks();
        try {
            Object bookmark = bookmarks.getByName(bookmarkName);
            XTextContent content = UnoRuntime.queryInterface(XTextContent.class, bookmark);
            XTextRange range = content.getAnchor();
            return range;
        } catch (NoSuchElementException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.WARNING, "unable to find "+bookmarkName+" bookmark");
        } catch (WrappedTargetException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.WARNING, "unable to find "+bookmarkName+" bookmark");
        }
        return null;
    }

    private XTextSection getSection(String sectionName, XComponent component){
        XTextSectionsSupplier supplier = UnoRuntime.queryInterface(XTextSectionsSupplier.class, component);
        XNameAccess sections = supplier.getTextSections();
        try {
            Object section = sections.getByName(sectionName);
            XTextSection textSection = UnoRuntime.queryInterface(XTextSection.class, section);
            return textSection;
        } catch (NoSuchElementException e) {
            Logger.getLogger(Ooops.class.getName()).log(Level.WARNING, "unable to find "+sectionName+" section");
        } catch (WrappedTargetException e) {
            Logger.getLogger(Ooops.class.getName()).log(Level.WARNING, "unable to find "+sectionName+" section");
        }
        return null;
    }

    private XTextTable getTextTable(String tableName, XComponent component){
        XTextTablesSupplier supplier = UnoRuntime.queryInterface(XTextTablesSupplier.class, component);
        XNameAccess tables = supplier.getTextTables();
        try {
            Object table = tables.getByName(tableName);
            XTextTable textTable = UnoRuntime.queryInterface(XTextTable.class, table);
            return textTable;
        } catch (NoSuchElementException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.WARNING, "unable to find " + tableName + " table");
        } catch (WrappedTargetException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.WARNING, "unable to find " + tableName + " table");
        }
        return null;
    }


    /*
     *   outs
     */
    public Ooops filter(Filter oooFilter){
        filter = oooFilter;
        return this;
    }

    public OutputStream stream(){
        if( component == null ){
            blank();
        }
        OoopsOutputStream outStream = new OoopsOutputStream();
        PropertyValue[] propertyValue = null;
        if( "writer_pdf_Export".equals(filter) ){
            propertyValue = new PropertyValue[3];
            propertyValue[2] = new PropertyValue();
            propertyValue[2].Name = "SelectPdfVersion";
            propertyValue[2].Value = 1; // PDF/A
        } else {
            propertyValue = new PropertyValue[2];
        }
        propertyValue[0] = new PropertyValue();
        propertyValue[0].Name = "OutputStream";
        propertyValue[0].Value = outStream;
        propertyValue[1] = new PropertyValue();
        propertyValue[1].Name = "FilterName";
        propertyValue[1].Value = filter.filterName();
        XStorable xstorable = UnoRuntime.queryInterface(XStorable.class, component);
        try {
            xstorable.storeToURL("private:stream", propertyValue);
            return outStream;
        } catch (com.sun.star.io.IOException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, null, ex);
        }
        return null;
    }

    public Boolean toUrl(String storeUrl){
        if( component == null ){
            blank();
        }
        PropertyValue[] propertyValue = new PropertyValue[1];
        propertyValue[0] = new PropertyValue();
        propertyValue[0].Name = "FilterName";
        propertyValue[0].Value = "writer_pdf_Export";
        XStorable xstorable = UnoRuntime.queryInterface(XStorable.class, component);
        try {
            xstorable.storeToURL(storeUrl, propertyValue);
            return true;
        } catch (com.sun.star.io.IOException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, null, ex);
        }
        return null;
    }

    public Boolean toStream(OutputStream outputStream){
        try {
            outputStream.write(toByteArray());
            return Boolean.TRUE;
        } catch (java.io.IOException e) {
            e.printStackTrace();
        }
        return Boolean.FALSE;
    }

    public byte[] toByteArray() {
        OoopsOutputStream stream = (OoopsOutputStream) stream();
        return stream.toByteArray();
    }


    private void disposeComponent() {
        XCloseable closeable = UnoRuntime.queryInterface(XCloseable.class, component);
        if (closeable != null) {
            try {
                closeable.close(true);
            } catch (CloseVetoException e) {
            }

        } else {
            component.dispose();
        }
    }




}
