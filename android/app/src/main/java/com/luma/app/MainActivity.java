package com.luma.app;

import android.app.Activity;
import android.content.ActivityNotFoundException;
import android.content.ClipData;
import android.content.Intent;
import android.content.res.Configuration;
import android.database.Cursor;
import android.graphics.Color;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.provider.OpenableColumns;
import android.view.View;
import android.view.WindowManager;
import android.webkit.JavascriptInterface;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceResponse;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.FrameLayout;
import android.widget.Toast;
import androidx.core.content.FileProvider;
import androidx.core.graphics.Insets;
import androidx.core.view.ViewCompat;
import androidx.core.view.WindowCompat;
import androidx.core.view.WindowInsetsCompat;
import androidx.core.view.WindowInsetsControllerCompat;
import androidx.documentfile.provider.DocumentFile;
import androidx.webkit.WebViewAssetLoader;
import com.chaquo.python.Python;
import com.chaquo.python.android.AndroidPlatform;
import org.json.JSONArray;
import org.json.JSONObject;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.UUID;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public final class MainActivity extends Activity {
    private static final int PICK_FILES = 11, PICK_FOLDER = 12, SAVE_FILE = 13;
    private static final long MAX_FILE = 32L * 1024 * 1024, MAX_TOTAL = 128L * 1024 * 1024;
    private static final String ORIGIN = "https://appassets.androidplatform.net";
    private static final Set<String> SHEETS = new HashSet<>(Arrays.asList("xlsx","xlsm","xltx","xltm","xls","xlt","xlsb","ods","ots","fods","xml"));
    private static final Set<String> DOCS = new HashSet<>(Arrays.asList("docx","docm","dotx","dotm","odt","ott","fodt","sxw","txt","text","html","htm","xhtml","rtf"));
    private final ExecutorService worker = Executors.newSingleThreadExecutor();
    private final AtomicBoolean busy = new AtomicBoolean(false), cancelled = new AtomicBoolean(false);
    private final ArrayList<JSONObject> selection = new ArrayList<>();
    private WebView browser;
    private FrameLayout container;
    private volatile boolean ready = false;
    private volatile String themeMode = "system";
    private JSONObject active, options = new JSONObject();
    private File inputs, sessions, pendingExport;
    private boolean pickerOpen = false;

    @Override public void onCreate(Bundle state) {
        super.onCreate(state);
        themeMode = getSharedPreferences("luma", MODE_PRIVATE).getString("theme", "system");
        try { options = new JSONObject(getSharedPreferences("luma", MODE_PRIVATE).getString("options", "{}")); } catch (Exception ignored) {}
        inputs = new File(getFilesDir(), "inputs");
        sessions = new File(getFilesDir(), "sessions");
        WindowCompat.setDecorFitsSystemWindows(getWindow(), false);
        container = new FrameLayout(this);
        ViewCompat.setOnApplyWindowInsetsListener(container, (view, insets) -> {
            int types = WindowInsetsCompat.Type.systemBars() | WindowInsetsCompat.Type.displayCutout();
            Insets bars = insets.getInsets(types), keyboard = insets.getInsets(WindowInsetsCompat.Type.ime());
            view.setPadding(bars.left, bars.top, bars.right, Math.max(bars.bottom, keyboard.bottom));
            return new WindowInsetsCompat.Builder(insets).setInsets(types | WindowInsetsCompat.Type.ime(), Insets.NONE)
                    .setInsetsIgnoringVisibility(types, Insets.NONE).setDisplayCutout(null).build();
        });
        browser = new WebView(this);
        browser.setId(R.id.browser);
        WebSettings settings = browser.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(false);
        settings.setAllowFileAccess(false);
        settings.setAllowContentAccess(false);
        settings.setMixedContentMode(WebSettings.MIXED_CONTENT_NEVER_ALLOW);
        settings.setSupportMultipleWindows(false);
        if (Build.VERSION.SDK_INT >= 33) settings.setAlgorithmicDarkeningAllowed(false);
        browser.addJavascriptInterface(new Bridge(), "LumaNative");
        WebViewAssetLoader loader = new WebViewAssetLoader.Builder()
                .addPathHandler("/assets/", new WebViewAssetLoader.AssetsPathHandler(this)).build();
        browser.setWebViewClient(new WebViewClient() {
            @Override public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
                return !isLocal(request.getUrl());
            }
            @Override public WebResourceResponse shouldInterceptRequest(WebView view, WebResourceRequest request) {
                if (isLocal(request.getUrl())) return loader.shouldInterceptRequest(request.getUrl());
                return new WebResourceResponse("text/plain", "UTF-8", new ByteArrayInputStream(new byte[0]));
            }
            @Override public void onPageFinished(WebView view, String url) { execute(MainActivity.this::sendState); }
        });
        container.addView(browser);
        setContentView(container);
        applyTheme();
        browser.loadUrl(ORIGIN + "/assets/index.html");
        execute(() -> {
            deleteTree(inputs); inputs.mkdirs(); sessions.mkdirs();
            recoverInterrupted();
            if (!Python.isStarted()) Python.start(new AndroidPlatform(getApplicationContext()));
            Python.getInstance().getModule("luma_engine");
            ready = true;
            sendState();
            importSharedIntent(getIntent());
        });
    }

    private boolean isLocal(Uri uri) { return "https".equals(uri.getScheme()) && "appassets.androidplatform.net".equals(uri.getHost()); }
    private boolean dark() { return "dark".equals(themeMode) || ("system".equals(themeMode) && (getResources().getConfiguration().uiMode & Configuration.UI_MODE_NIGHT_MASK) == Configuration.UI_MODE_NIGHT_YES); }
    private void applyTheme() {
        boolean dark = dark(); int color = Color.parseColor(dark ? "#141320" : "#F4F3FA");
        container.setBackgroundColor(color); browser.setBackgroundColor(color);
        getWindow().setStatusBarColor(color); getWindow().setNavigationBarColor(color);
        if (Build.VERSION.SDK_INT >= 29) { getWindow().setStatusBarContrastEnforced(false); getWindow().setNavigationBarContrastEnforced(false); }
        WindowInsetsControllerCompat controller = WindowCompat.getInsetsController(getWindow(), container);
        controller.setAppearanceLightStatusBars(!dark); controller.setAppearanceLightNavigationBars(!dark);
    }
    @Override public void onConfigurationChanged(Configuration config) {
        super.onConfigurationChanged(config); applyTheme();
        emit("theme", json("mode",themeMode,"resolved",dark()?"dark":"light"));
    }
    private final class Bridge {
        @JavascriptInterface public String theme() { return dark()?"dark":"light"; }
        @JavascriptInterface public void request(String action, String payload) {
            if (payload == null || payload.length() > 10000) return;
            try {
                JSONObject data = new JSONObject(payload);
                if ("theme".equals(action)) {
                    String mode = data.optString("mode");
                    if (!Arrays.asList("light","dark","system").contains(mode)) return;
                    runOnUiThread(() -> { themeMode=mode; getSharedPreferences("luma", MODE_PRIVATE).edit().putString("theme",mode).apply(); applyTheme(); emit("theme",json("mode",mode,"resolved",dark()?"dark":"light")); });
                } else if ("cancel".equals(action)) cancelled.set(true);
                else if ("pickFiles".equals(action) || "pickFolder".equals(action)) runOnUiThread(() -> pick("pickFolder".equals(action)));
                else execute(() -> handle(action, data));
            } catch (Exception e) { error("No se pudo leer la solicitud. Vuelve a intentarlo."); }
        }
    }
    @FunctionalInterface private interface Task { void run() throws Exception; }
    private void execute(Task task) {
        if (worker.isShutdown()) return;
        worker.execute(() -> {
            try { task.run(); }
            catch (Exception e) { error(friendly(e)); }
        });
    }
    private void handle(String action, JSONObject data) throws Exception {
        if ("state".equals(action)) { sendState(); return; }
        if ("file".equals(action)) { fileAction(data); return; }
        if (busy.get() || !ready) { toast("Espera a que termine la operación actual."); return; }
        switch (action) {
            case "options": saveOptions(data); break;
            case "remove": {
                String id=data.optString("id");
                for (int i=selection.size()-1;i>=0;i--) if (id.equals(selection.get(i).optString("id"))) {
                    deleteTree(new File(inputs,id)); selection.remove(i);
                }
                sendState(); break;
            }
            case "clearSelection": selection.clear(); deleteTree(inputs); inputs.mkdirs(); sendState(); break;
            case "demo": demo(); break;
            case "convert": convert(data); break;
            case "delete": {
                String id=data.optString("id");
                if ("all".equals(id)) { deleteTree(sessions); sessions.mkdirs(); }
                else deleteTree(sessionDir(id));
                sendState(); toast("Resultados eliminados de Luma."); break;
            }
        }
    }
    private void saveOptions(JSONObject data) {
        options=data;
        getSharedPreferences("luma",MODE_PRIVATE).edit().putString("options",data.toString()).apply();
    }
    private void pick(boolean folder) {
        if (busy.get() || !ready || pickerOpen) return;
        Intent intent = new Intent(folder ? Intent.ACTION_OPEN_DOCUMENT_TREE : Intent.ACTION_OPEN_DOCUMENT);
        if (!folder) { intent.addCategory(Intent.CATEGORY_OPENABLE); intent.setType("*/*"); intent.putExtra(Intent.EXTRA_ALLOW_MULTIPLE,true); }
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        try { pickerOpen=true; startActivityForResult(intent, folder?PICK_FOLDER:PICK_FILES); }
        catch (ActivityNotFoundException e) { pickerOpen=false; error("No se encontró un selector de archivos en este teléfono."); }
    }
    @Override protected void onActivityResult(int request, int result, Intent data) {
        super.onActivityResult(request,result,data);
        pickerOpen=false;
        if (request==SAVE_FILE) {
            File source=pendingExport; pendingExport=null;
            if (result==RESULT_OK && data!=null && data.getData()!=null && source!=null) {
                Uri target=data.getData();
                execute(() -> {
                    try (InputStream in=new FileInputStream(source); OutputStream out=getContentResolver().openOutputStream(target,"wt")) {
                        if (out==null) throw new Exception("No se pudo abrir el destino.");
                        copy(in,out,Long.MAX_VALUE);
                    } catch (Exception e) {
                        try { android.provider.DocumentsContract.deleteDocument(getContentResolver(),target); } catch (Exception ignored) {}
                        throw e;
                    }
                    toast("Archivo guardado en la ubicación que elegiste.");
                });
            }
            return;
        }
        if (result!=RESULT_OK || data==null) return;
        if (request==PICK_FILES) {
            List<Uri> uris=new ArrayList<>();
            if (data.getClipData()!=null) for (int i=0;i<data.getClipData().getItemCount();i++) uris.add(data.getClipData().getItemAt(i).getUri());
            else if (data.getData()!=null) uris.add(data.getData());
            importUris(uris);
        } else if (request==PICK_FOLDER && data.getData()!=null) {
            Uri uri=data.getData();
            execute(() -> importFolder(uri));
        }
    }
    private void importSharedIntent(Intent intent) {
        if (intent==null) return;
        List<Uri> uris=new ArrayList<>();
        if (Intent.ACTION_SEND.equals(intent.getAction())) {
            Uri uri=intent.getParcelableExtra(Intent.EXTRA_STREAM); if(uri!=null)uris.add(uri);
        } else if(Intent.ACTION_SEND_MULTIPLE.equals(intent.getAction())) {
            ArrayList<Uri> list=intent.getParcelableArrayListExtra(Intent.EXTRA_STREAM); if(list!=null)uris.addAll(list);
        }
        if(!uris.isEmpty())importUris(uris);
    }
    @Override protected void onNewIntent(Intent intent){super.onNewIntent(intent);setIntent(intent);execute(()->importSharedIntent(intent));}
    private void importUris(List<Uri> uris) {
        execute(() -> {
            if(!busy.compareAndSet(false,true))return;
            active=json("label","Añadiendo archivos"); sendState();
            ArrayList<String> issues=new ArrayList<>(); int added=0;
            try {
                for(Uri uri:uris) {
                    try { String name=displayName(uri); importOne(uri,name,name); added++; }
                    catch(Exception e){issues.add(friendly(e));}
                }
            } finally {busy.set(false);active=null;sendState();}
            importNotice(added,issues);
        });
    }
    private void importFolder(Uri uri) throws Exception {
        if(!busy.compareAndSet(false,true))return;
        active=json("label","Explorando carpeta"); sendState();
        ArrayList<String> issues=new ArrayList<>(); int before=selection.size(); int[] visited={0};
        try {
            DocumentFile folder=DocumentFile.fromTreeUri(this,uri);
            if(folder==null || !folder.canRead())throw new Exception("No se pudo leer la carpeta seleccionada.");
            walk(folder,"",0,visited,issues);
        } finally {busy.set(false);active=null;sendState();}
        importNotice(selection.size()-before,issues);
    }
    private void walk(DocumentFile folder,String prefix,int depth,int[] visited,List<String> issues) throws Exception {
        if(depth>20)throw new Exception("La carpeta tiene demasiados niveles. Elige una subcarpeta.");
        for(DocumentFile file:folder.listFiles()) {
            if(++visited[0]>3000)throw new Exception("La carpeta contiene demasiados elementos. Elige una subcarpeta.");
            String name=file.getName(); if(name==null)continue;
            if(name.startsWith("~$")||name.startsWith(".~lock.")||name.startsWith("CONVERSIONES_OFFICE"))continue;
            String relative=prefix+name;
            if(file.isDirectory())walk(file,relative+"/",depth+1,visited,issues);
            else if(file.isFile()&&supported(name)) {
                if(selection.size()>=100)throw new Exception("Llegaste al límite de 100 archivos. Convierte este lote primero.");
                try{importOne(file.getUri(),name,relative);}catch(Exception e){issues.add(name+": "+friendly(e));}
            }
        }
    }
    private void importNotice(int added,List<String> issues) {
        if(!issues.isEmpty())toast(added+" archivos añadidos. "+issues.get(0)+(issues.size()>1?" (y "+(issues.size()-1)+" más)":""));
        else toast(added>0?added+" "+(added==1?"archivo añadido.":"archivos añadidos."):"No encontramos formatos compatibles en esa selección.");
    }
    private String displayName(Uri uri) throws Exception {
        if(!"content".equals(uri.getScheme()))throw new Exception("El archivo debe venir del selector de Android.");
        try(Cursor cursor=getContentResolver().query(uri,new String[]{OpenableColumns.DISPLAY_NAME},null,null,null)) {
            if(cursor!=null&&cursor.moveToFirst())return cursor.getString(0);
        }
        throw new Exception("No se pudo obtener el nombre del archivo.");
    }
    private static String ext(String name) {int pos=name.lastIndexOf('.');return pos<0?"":name.substring(pos+1).toLowerCase(Locale.ROOT);}
    private static boolean supported(String name){return SHEETS.contains(ext(name))||DOCS.contains(ext(name));}
    private static String cleanName(String name) {
        String clean=name.replaceAll("[\\\\/:*?\"<>|\\p{Cntrl}]","_").trim();
        if(clean.isEmpty()||clean.equals(".")||clean.equals(".."))clean="archivo";
        if(clean.length()>140){String suffix=ext(clean);clean=clean.substring(0,120)+"."+suffix;}
        return clean;
    }
    private long selectedBytes(){long sum=0;for(JSONObject file:selection)sum+=file.optLong("bytes");return sum;}
    private void importOne(Uri uri,String name,String relative) throws Exception {
        if(name==null||!supported(name))throw new Exception("Formato no compatible: "+name+". Exporta a XLSX, DOCX, ODT o TXT.");
        if(selection.size()>=100)throw new Exception("Máximo 100 archivos por lote.");
        String id=UUID.randomUUID().toString(); File folder=new File(inputs,id); folder.mkdirs();
        File target=new File(folder,cleanName(name));
        long remaining=MAX_TOTAL-selectedBytes();
        try(InputStream in=getContentResolver().openInputStream(uri);OutputStream out=new FileOutputStream(target)) {
            if(in==null)throw new Exception("No se pudo abrir "+name+".");
            long bytes=copy(in,out,Math.min(MAX_FILE,remaining));
            if(bytes==0)throw new Exception("El archivo "+name+" está vacío.");
            selection.add(json("id",id,"name",target.getName(),"relative",relative,"bytes",bytes,"path",target.getAbsolutePath()));
        }catch(Exception e){deleteTree(folder);throw e;}
    }
    private static long copy(InputStream in,OutputStream out,long limit)throws Exception {
        byte[] buffer=new byte[32768];long bytes=0;int count;
        while((count=in.read(buffer))!=-1){bytes+=count;if(bytes>limit)throw new Exception("El archivo supera 32 MB o el lote supera 128 MB.");out.write(buffer,0,count);}
        return bytes;
    }
    private void demo() throws Exception {
        File temp=new File(getCacheDir(),"demo");temp.mkdirs();
        Python.getInstance().getModule("luma_engine").callAttr("create_demo",temp.getAbsolutePath());
        for(String name:Arrays.asList("Bienvenida.txt","Pequeño inventario.xlsx")) {
            if(selection.size()>=100)break;
            String id=UUID.randomUUID().toString();File dir=new File(inputs,id);dir.mkdirs();File source=new File(temp,name),target=new File(dir,name);
            try(InputStream in=new FileInputStream(source);OutputStream out=new FileOutputStream(target)){copy(in,out,MAX_FILE);}
            selection.add(json("id",id,"name",name,"relative",name,"bytes",target.length(),"path",target.getAbsolutePath()));
        }
        deleteTree(temp);sendState();toast("Dos archivos de ejemplo, listos para probar.");
    }
    private void convert(JSONObject settings)throws Exception {
        String kind=settings.optString("kind","both");
        if(!Arrays.asList("both","excel","docs").contains(kind))throw new Exception("El tipo de conversión no es válido.");
        List<JSONObject> files=new ArrayList<>();
        for(JSONObject file:selection)if("both".equals(kind)||("excel".equals(kind)==SHEETS.contains(ext(file.optString("name")))))files.add(file);
        if(files.isEmpty())throw new Exception("Añade archivos del tipo que elegiste.");
        if(!busy.compareAndSet(false,true))return;
        cancelled.set(false);saveOptions(settings);
        String id=UUID.randomUUID().toString();File dir=sessionDir(id);dir.mkdirs();
        JSONObject job=json("id",id,"time",System.currentTimeMillis(),"total",files.size(),"processed",0,"outputCount",0,"errors",0,"warnings",false,"running",true,"results",new JSONArray(),"outputs",new JSONArray());
        active=json("label","Transformando archivos","total",files.size(),"processed",0);
        runOnUiThread(()->getWindow().addFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON));
        try {
            writeJob(dir,job);sendState();
            for(JSONObject input:files) {
                if(cancelled.get()){job.put("cancelled",true);break;}
                String name=input.getString("name");active.put("current",name);sendState();
                File staging=new File(dir,"staging");staging.mkdirs();
                JSONObject record=json("name",input.optString("relative",name),"outputs",new JSONArray(),"warnings",new JSONArray());
                try {
                    JSONObject response=new JSONObject(Python.getInstance().getModule("luma_engine").callAttr("convert",input.getString("path"),staging.getAbsolutePath(),settings.toString()).toString());
                    JSONArray names=response.optJSONArray("files");
                    if("pdf".equals(response.getString("kind"))) {
                        String filename=cleanName(name.substring(0,name.lastIndexOf('.')))+".pdf";
                        PdfExporter.write(response.getJSONArray("lines"),name,new File(staging,filename));
                        names=new JSONArray().put(filename);
                    }
                    if(names==null||names.length()==0)throw new Exception("No se generaron resultados.");
                    ArrayList<File> moved=new ArrayList<>();
                    try {
                        for(int n=0;n<names.length();n++) {
                            String filename=names.getString(n);File source=new File(staging,filename);File target=unique(dir,filename);
                            if(!source.isFile()||source.length()==0||!source.renameTo(target))throw new Exception("No se pudo conservar el archivo de salida.");
                            moved.add(target);record.getJSONArray("outputs").put(target.getName());
                        }
                    }catch(Exception e){for(File file:moved)file.delete();throw e;}
                    for(File file:moved)job.getJSONArray("outputs").put(file.getName());
                    JSONArray warnings=response.optJSONArray("warnings");if(warnings!=null)record.put("warnings",warnings);
                    if(warnings!=null&&warnings.length()>0)job.put("warnings",true);
                }catch(Exception e){record.put("outputs",new JSONArray());record.put("error",friendly(e));job.put("errors",job.optInt("errors")+1);}
                finally{deleteTree(staging);}
                job.getJSONArray("results").put(record);job.put("processed",job.optInt("processed")+1);job.put("outputCount",job.getJSONArray("outputs").length());
                writeJob(dir,job);active.put("processed",job.optInt("processed"));sendState();
            }
            job.put("running",false);writeJob(dir,job);packageJob(dir,job);
            // Libera únicamente los originales que se procesaron; el resto sigue en la selección.
            int processed=job.optInt("processed");
            for(int i=0;i<processed;i++){JSONObject item=files.get(i);selection.remove(item);deleteTree(new File(inputs,item.getString("id")));}
        }catch(Exception e){job.put("running",false);job.put("interrupted",true);writeJob(dir,job);throw e;}
        finally{busy.set(false);active=null;runOnUiThread(()->getWindow().clearFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON));sendState();}
        emit("done",json("id",id));
    }
    private static File unique(File dir,String name){File file=new File(dir,name);int n=2;int pos=name.lastIndexOf('.');String base=pos>0?name.substring(0,pos):name,suffix=pos>0?name.substring(pos):"";while(file.exists()||Arrays.asList("Luma-reporte.csv","Luma-resultados.zip","session.json").contains(file.getName()))file=new File(dir,base+"_"+(n++)+suffix);return file;}
    private File sessionDir(String id)throws Exception {
        if(!id.matches("[0-9a-f-]{36}"))throw new Exception("No se encontró la conversión.");
        File result=new File(sessions,id);
        if(!result.getCanonicalFile().getParentFile().equals(sessions.getCanonicalFile()))throw new Exception("Ruta no válida.");
        return result;
    }
    private static void writeJob(File dir,JSONObject job)throws Exception {
        File temp=new File(dir,"session.part");
        try(FileOutputStream out=new FileOutputStream(temp)){out.write(job.toString().getBytes(StandardCharsets.UTF_8));out.getFD().sync();}
        if(!temp.renameTo(new File(dir,"session.json")))throw new Exception("No se pudo guardar el historial.");
    }
    private JSONObject readJob(File dir)throws Exception{
        try(InputStream in=new FileInputStream(new File(dir,"session.json"));ByteArrayOutputStream out=new ByteArrayOutputStream()){
            copy(in,out,8*1024*1024);return new JSONObject(out.toString("UTF-8"));
        }
    }
    private JSONArray history() {
        ArrayList<JSONObject> list=new ArrayList<>();File[] dirs=sessions.listFiles();
        if(dirs!=null)for(File dir:dirs)try{JSONObject job=readJob(dir);if(!job.optBoolean("running"))list.add(job);}catch(Exception ignored){}
        list.sort(Comparator.comparingLong((JSONObject j)->j.optLong("time")).reversed());
        return new JSONArray(list);
    }
    private void recoverInterrupted()throws Exception {
        File[] dirs=sessions.listFiles();if(dirs==null)return;
        for(File dir:dirs) {
            if(!dir.isDirectory())continue;
            deleteTree(new File(dir,"staging"));
            try {
                JSONObject job=readJob(dir);
                if(job.optBoolean("running")){job.put("running",false);job.put("interrupted",true);writeJob(dir,job);}
                if(!new File(dir,"Luma-resultados.zip").exists())packageJob(dir,job);
            }catch(Exception ignored){}
        }
    }
    private static String csv(String value){return "\""+value.replace("\"","\"\"")+"\"";}
    private void packageJob(File dir,JSONObject job)throws Exception {
        StringBuilder report=new StringBuilder("\ufeffArchivo,Estado,Resultados,Avisos,Error\n");
        JSONArray results=job.getJSONArray("results");
        for(int i=0;i<results.length();i++) {
            JSONObject row=results.getJSONObject(i);
            report.append(csv(reportSafe(row.optString("name")))).append(',').append(csv(row.has("error")?"ERROR":"OK")).append(',').append(csv(row.optJSONArray("outputs").join(" | "))).append(',').append(csv(reportSafe(row.optJSONArray("warnings").join(" | ")))).append(',').append(csv(reportSafe(row.optString("error")))).append('\n');
        }
        File reportFile=new File(dir,"Luma-reporte.csv");
        try(FileOutputStream out=new FileOutputStream(reportFile)){out.write(report.toString().getBytes(StandardCharsets.UTF_8));}
        File temp=new File(dir,"Luma-resultados.part");
        try(ZipOutputStream zip=new ZipOutputStream(new FileOutputStream(temp))) {
            ArrayList<String> names=new ArrayList<>();JSONArray outputs=job.getJSONArray("outputs");for(int i=0;i<outputs.length();i++)names.add(outputs.getString(i));names.add(reportFile.getName());
            for(String name:names){zip.putNextEntry(new ZipEntry(name));try(InputStream in=new FileInputStream(new File(dir,name))){copy(in,zip,Long.MAX_VALUE);}zip.closeEntry();}
        }
        if(!temp.renameTo(new File(dir,"Luma-resultados.zip")))throw new Exception("No se pudo preparar el ZIP.");
    }
    private static String reportSafe(String value){return value.matches("^[=+@\\-\\t\\r\\n].*")?"'"+value:value;}
    private void fileAction(JSONObject data)throws Exception {
        File dir=sessionDir(data.getString("session"));JSONObject job=readJob(dir);String name=data.getString("name");
        if("@zip".equals(name))name="Luma-resultados.zip";
        else if("@report".equals(name))name="Luma-reporte.csv";
        else {boolean found=false;JSONArray outputs=job.getJSONArray("outputs");for(int i=0;i<outputs.length();i++)if(name.equals(outputs.getString(i)))found=true;if(!found)throw new Exception("El resultado ya no está disponible.");}
        File file=new File(dir,name);
        if(!file.getCanonicalFile().getParentFile().equals(dir.getCanonicalFile())||!file.isFile())throw new Exception("El resultado ya no está disponible.");
        String type=name.endsWith(".pdf")?"application/pdf":name.endsWith(".zip")?"application/zip":"text/csv";
        String action=data.getString("action");
        runOnUiThread(() -> {
            try {
                if("save".equals(action)) {
                    if(pickerOpen){toast("Termina de elegir la ubicación primero.");return;}
                    pendingExport=file;Intent intent=new Intent(Intent.ACTION_CREATE_DOCUMENT);intent.addCategory(Intent.CATEGORY_OPENABLE);intent.setType(type);intent.putExtra(Intent.EXTRA_TITLE,file.getName());pickerOpen=true;startActivityForResult(intent,SAVE_FILE);
                } else {
                    Uri uri=FileProvider.getUriForFile(this,"com.luma.app.files",file);Intent intent;
                    if("share".equals(action)){intent=new Intent(Intent.ACTION_SEND);intent.setType(type);intent.putExtra(Intent.EXTRA_STREAM,uri);}
                    else if("open".equals(action)){intent=new Intent(Intent.ACTION_VIEW);intent.setDataAndType(uri,type);}else return;
                    intent.setClipData(ClipData.newRawUri(file.getName(),uri));intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                    startActivity(Intent.createChooser(intent,"share".equals(action)?"Compartir con":"Abrir con"));
                }
            }catch(ActivityNotFoundException e){pickerOpen=false;error("Instala una app que pueda abrir este formato o guárdalo para usarlo después.");}
            catch(Exception e){pickerOpen=false;error("No se pudo abrir esa acción. Intenta guardar el archivo.");}
        });
    }
    private void sendState() {
        JSONObject value=json("ready",ready,"busy",busy.get(),"selection",new JSONArray(selection),"history",history(),"active",active==null?JSONObject.NULL:active,"options",options,"themeMode",themeMode,"theme",dark()?"dark":"light");
        emit("state",value);
    }
    private void emit(String type,JSONObject value) {
        runOnUiThread(() -> {if(!isDestroyed()&&browser!=null)browser.evaluateJavascript("window.Luma && window.Luma.onEvent("+JSONObject.quote(type)+","+value+")",null);});
    }
    private void toast(String message){emit("toast",json("message",message));}
    private void error(String message){emit("error",json("message",message));}
    private static String friendly(Exception error) {
        String message=error.getMessage();if(message==null||message.trim().isEmpty())return "No se pudo procesar el archivo. Comprueba que esté completo y sin contraseña.";
        if(message.contains("BadZipFile")||message.contains("not a zip")||message.contains("InvalidFileException"))return "El documento está dañado o su extensión no corresponde al contenido. Guarda otra copia desde su aplicación original.";
        if(message.contains("UnicodeEncodeError"))return "Hay caracteres que Windows-1252 no puede guardar. Elige UTF-8 y vuelve a convertir.";
        if(message.contains("Permission denied"))return "No hay permiso para leer el archivo. Selecciónalo otra vez.";
        if(message.toLowerCase(Locale.ROOT).contains("encrypted")||message.toLowerCase(Locale.ROOT).contains("password"))return "El documento está cifrado. Guarda una copia sin contraseña y vuelve a intentarlo.";
        if(message.contains("No se pudo convertir el archivo Excel"))return "No se pudo leer la hoja. Comprueba que el archivo esté completo o guarda otra copia en formato XLSX u ODS.";
        if(message.length()>450)message=message.substring(0,450)+"…";
        return message.replaceFirst("^[A-Za-z_.]+(?:Error|Exception): ","");
    }
    private static JSONObject json(Object... pairs){JSONObject object=new JSONObject();try{for(int i=0;i<pairs.length;i+=2)object.put((String)pairs[i],pairs[i+1]);}catch(Exception ignored){}return object;}
    private static void deleteTree(File file){if(file==null||!file.exists())return;if(file.isDirectory()){File[] children=file.listFiles();if(children!=null)for(File child:children)deleteTree(child);}file.delete();}
    @Override public void onBackPressed() {
        browser.evaluateJavascript("window.Luma && window.Luma.back()",result->{if(!"true".equals(result)){if(busy.get())Toast.makeText(this,"Puedes detener el lote desde la pantalla de conversión.",Toast.LENGTH_SHORT).show();else super.onBackPressed();}});
    }
    @Override protected void onDestroy(){super.onDestroy();cancelled.set(true);worker.shutdown();if(browser!=null){browser.removeJavascriptInterface("LumaNative");browser.destroy();browser=null;}}
}
