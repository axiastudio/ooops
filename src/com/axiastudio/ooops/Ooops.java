package com.axiastudio.ooops;

import com.axiastudio.ooops.filters.Filter;
import com.axiastudio.ooops.streams.OoopsInputStream;
import com.axiastudio.ooops.streams.OoopsOutputStream;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.*;
import com.sun.star.text.XBookmarksSupplier;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.OutputStream;
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
    private Filter filter = Filter.writer8;

    public static Ooops create() {
        return new Ooops();
    }

    public Ooops open(String connectionString){
        try{
            XComponentContext context = com.sun.star.comp.helper.Bootstrap.createInitialComponentContext(null);
            XMultiComponentFactory factory = context.getServiceManager();
            Object objResolver = factory.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", context);
            XUnoUrlResolver resolver = UnoRuntime.queryInterface(XUnoUrlResolver.class, objResolver);
            Object objectInitial = resolver.resolve(connectionString);
            factory = UnoRuntime.queryInterface(XMultiComponentFactory.class, objectInitial);
            XPropertySet properties = UnoRuntime.queryInterface(XPropertySet.class, factory);
            Object objContext = properties.getPropertyValue("DefaultContext");
            context = UnoRuntime.queryInterface(XComponentContext.class, objContext);
            loader = UnoRuntime.queryInterface(XComponentLoader.class, factory.createInstanceWithContext("com.sun.star.frame.Desktop", context));
        } catch (Exception ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, null, ex);
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
        component = loadDocumentComponent(ooopsInputStream);
        return this;
    }

    public Ooops fromMap(Map<String, Object> values){
        for( String key: values.keySet() ){
            XTextRange anchor = this.getAnchor(key, component);
            if( anchor != null ){
                Object value = values.get(key);
                if( value != null ){
                    anchor.setString(value.toString());
                }
            }
        }
        return this;
    }

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
        propertyValue[1].Value = filter.name();
        XStorable xstorable = UnoRuntime.queryInterface(XStorable.class, component);
        try {
            xstorable.storeToURL("private:stream", propertyValue);
            return outStream;
        } catch (IOException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, null, ex);
        }
        return null;
    }

    public Boolean toStream(OutputStream outputStream){
        OoopsOutputStream stream = (OoopsOutputStream) stream();
        try {
            outputStream.write(stream.toByteArray());
            return Boolean.TRUE;
        } catch (java.io.IOException e) {
            e.printStackTrace();
        }
        return Boolean.FALSE;
    }



    private XComponent loadDocumentComponent(OoopsInputStream inStream){
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
                XComponent xComponent = loader.loadComponentFromURL("private:stream", "_blank", 0, propertyValue);
                return xComponent;
            } catch (com.sun.star.lang.IllegalArgumentException e) {
                e.printStackTrace();
            }
        } catch (IOException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.SEVERE, null, ex);
        }
        return null;
    }

    private XTextRange getAnchor(String anchorName, XComponent component){
        XBookmarksSupplier supplier = UnoRuntime.queryInterface(XBookmarksSupplier.class, component);
        XNameAccess bookmarks = supplier.getBookmarks();
        try {
            Object myBookmark = bookmarks.getByName(anchorName);
            XTextContent content = UnoRuntime.queryInterface(XTextContent.class, myBookmark);
            XTextRange range = content.getAnchor();
            return range;
        } catch (NoSuchElementException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.WARNING, "unable to find "+anchorName+" anchor", ex);
        } catch (WrappedTargetException ex) {
            Logger.getLogger(Ooops.class.getName()).log(Level.WARNING, "unable to find "+anchorName+" anchor", ex);
        }
        return null;
    }




}
