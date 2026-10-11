package com.luma.app;

import android.content.ContentProvider;
import android.content.ContentValues;
import android.database.Cursor;
import android.database.MatrixCursor;
import android.net.Uri;
import android.os.ParcelFileDescriptor;
import android.provider.OpenableColumns;
import java.io.File;
import java.io.FileNotFoundException;

/** Proveedor que ejercita ContentResolver sin depender de un servicio de nube. */
public class TestDocuments extends ContentProvider {
    @Override public boolean onCreate(){return true;}
    private File file(Uri uri)throws FileNotFoundException {
        String name=uri.getLastPathSegment();
        if(name==null||!name.matches("[a-zA-Z0-9_.-]+"))throw new FileNotFoundException();
        return new File(getContext().getCacheDir(),name);
    }
    @Override public ParcelFileDescriptor openFile(Uri uri,String mode)throws FileNotFoundException{return ParcelFileDescriptor.open(file(uri),ParcelFileDescriptor.parseMode(mode));}
    @Override public String getType(Uri uri){return uri.toString().endsWith(".txt")?"text/plain":"application/pdf";}
    @Override public Cursor query(Uri uri,String[] projection,String selection,String[] args,String order) {
        String[] columns=projection!=null?projection:new String[]{OpenableColumns.DISPLAY_NAME,OpenableColumns.SIZE};
        MatrixCursor cursor=new MatrixCursor(columns);Object[] row=new Object[columns.length];
        for(int i=0;i<columns.length;i++)row[i]=OpenableColumns.SIZE.equals(columns[i])?0:uri.getLastPathSegment();
        cursor.addRow(row);return cursor;
    }
    @Override public int delete(Uri uri,String selection,String[] args){try{return file(uri).delete()?1:0;}catch(Exception e){return 0;}}
    @Override public Uri insert(Uri uri,ContentValues values){throw new UnsupportedOperationException();}
    @Override public int update(Uri uri,ContentValues values,String selection,String[] args){throw new UnsupportedOperationException();}
}
