package com.luma.app;

import android.content.Context;
import android.graphics.Bitmap;
import android.graphics.Rect;
import android.graphics.pdf.PdfRenderer;
import android.os.ParcelFileDescriptor;
import android.webkit.WebView;
import androidx.core.graphics.Insets;
import androidx.core.view.ViewCompat;
import androidx.core.view.WindowCompat;
import androidx.core.view.WindowInsetsCompat;
import androidx.test.core.app.ActivityScenario;
import androidx.test.ext.junit.runners.AndroidJUnit4;
import androidx.test.platform.app.InstrumentationRegistry;
import com.chaquo.python.Python;
import org.json.JSONArray;
import org.json.JSONObject;
import org.junit.Test;
import org.junit.runner.RunWith;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;
import java.util.zip.ZipFile;
import static org.junit.Assert.*;

@RunWith(AndroidJUnit4.class)
public class LumaSmokeTest {
    private String evaluate(ActivityScenario<MainActivity> scenario,String script)throws Exception {
        CountDownLatch latch=new CountDownLatch(1);AtomicReference<String> answer=new AtomicReference<>("");
        scenario.onActivity(activity->((WebView)activity.findViewById(R.id.browser)).evaluateJavascript(script,value->{answer.set(value);latch.countDown();}));
        assertTrue("WebView no respondió",latch.await(15,TimeUnit.SECONDS));return answer.get();
    }
    private void waitFor(ActivityScenario<MainActivity> scenario,String condition)throws Exception {
        for(int i=0;i<120;i++){if("true".equals(evaluate(scenario,condition)))return;Thread.sleep(500);}fail("No se cumplió: "+condition);
    }
    private void screenshot(String name)throws Exception {
        InstrumentationRegistry.getInstrumentation().waitForIdleSync();Thread.sleep(700);
        Bitmap bitmap=InstrumentationRegistry.getInstrumentation().getUiAutomation().takeScreenshot();
        assertNotNull(bitmap);
        File target=new File(InstrumentationRegistry.getInstrumentation().getTargetContext().getExternalFilesDir(null),name);
        try(FileOutputStream out=new FileOutputStream(target)){assertTrue(bitmap.compress(Bitmap.CompressFormat.PNG,100,out));}
        bitmap.recycle();
    }
    private static String read(File file)throws Exception {
        try(InputStream in=new FileInputStream(file);ByteArrayOutputStream out=new ByteArrayOutputStream()){byte[] buffer=new byte[8192];int n;while((n=in.read(buffer))!=-1)out.write(buffer,0,n);return out.toString("UTF-8");}
    }

    @Test public void actualPythonEngineAndUnicodePdfWorkOnAndroid()throws Exception {
        try(ActivityScenario<MainActivity> scenario=ActivityScenario.launch(MainActivity.class)) {
            waitFor(scenario,"window.Luma && state.ready");
            Context context=InstrumentationRegistry.getInstrumentation().getTargetContext();
            File dir=new File(context.getCacheDir(),"engine-smoke");dir.mkdirs();
            JSONObject result=new JSONObject(Python.getInstance().getModule("luma_engine").callAttr("smoke",dir.getAbsolutePath()).toString());
            File pdf=new File(dir,"unicode.pdf");PdfExporter.write(result.getJSONArray("lines"),"España · Café",pdf);
            assertTrue(pdf.length()>1000);
            try(ParcelFileDescriptor descriptor=ParcelFileDescriptor.open(pdf,ParcelFileDescriptor.MODE_READ_ONLY);PdfRenderer renderer=new PdfRenderer(descriptor)) {
                assertTrue(renderer.getPageCount()>=1);PdfRenderer.Page page=renderer.openPage(0);
                Bitmap image=Bitmap.createBitmap(595,842,Bitmap.Config.ARGB_8888);page.render(image,null,null,PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY);page.close();image.recycle();
            }
            JSONArray longText=new JSONArray();for(int i=0;i<220;i++)longText.put("Línea "+i+": café, niña y té. Texto largo que continúa en varias páginas sin cortarse.");
            File multiple=new File(dir,"multiple.pdf");PdfExporter.write(longText,"Documento largo",multiple);
            try(ParcelFileDescriptor descriptor=ParcelFileDescriptor.open(multiple,ParcelFileDescriptor.MODE_READ_ONLY);PdfRenderer renderer=new PdfRenderer(descriptor)){assertTrue(renderer.getPageCount()>2);}
        }
    }

    @Test public void completeBatchHistoryAndThemesPersist()throws Exception {
        try(ActivityScenario<MainActivity> scenario=ActivityScenario.launch(MainActivity.class)) {
            waitFor(scenario,"window.Luma && state.ready");
            evaluate(scenario,"send('theme',{mode:'light'})");waitFor(scenario,"document.documentElement.dataset.theme==='light'");
            screenshot("luma-android-light.png");
            evaluate(scenario,"send('theme',{mode:'dark'})");waitFor(scenario,"document.documentElement.dataset.theme==='dark'");
            screenshot("luma-android-dark.png");
            scenario.onActivity(activity->{
                WebView view=activity.findViewById(R.id.browser);Insets bars=ViewCompat.getRootWindowInsets(activity.getWindow().getDecorView()).getInsets(WindowInsetsCompat.Type.systemBars()|WindowInsetsCompat.Type.displayCutout());
                int[] location=new int[2];view.getLocationOnScreen(location);android.util.DisplayMetrics display=new android.util.DisplayMetrics();activity.getWindowManager().getDefaultDisplay().getRealMetrics(display);
                assertEquals(display.heightPixels-bars.bottom,location[1]+view.getHeight(),2);
                assertFalse(WindowCompat.getInsetsController(activity.getWindow(),view).isAppearanceLightStatusBars());
            });
            assertEquals("true",evaluate(scenario,"Math.abs(innerHeight-document.querySelector('.bottom-nav').getBoundingClientRect().bottom-16)<2"));
            evaluate(scenario,"document.querySelector('#demo').click()");waitFor(scenario,"state.selection.length===2");
            evaluate(scenario,"document.querySelector('#convert').click()");
            waitFor(scenario,"state.history.length>0 && state.history[0].outputCount===2 && !state.busy && !document.querySelector('#modal-wrap').hidden");
            JSONArray sessions=new JSONArray(evaluate(scenario,"state.history"));JSONObject job=sessions.getJSONObject(0);assertEquals(0,job.getInt("errors"));
            File dir=new File(InstrumentationRegistry.getInstrumentation().getTargetContext().getFilesDir(),"sessions/"+job.getString("id"));
            assertTrue(new File(dir,"Luma-reporte.csv").isFile());
            try(ZipFile zip=new ZipFile(new File(dir,"Luma-resultados.zip"))){assertEquals(3,zip.size());assertNotNull(zip.getEntry("Bienvenida.pdf"));}
            assertTrue(read(new File(dir,"Pequeño inventario.csv")).contains("Café"));
            screenshot("luma-android-results.png");
            evaluate(scenario,"document.querySelector('.modal-close').click();document.querySelector('[data-page=history]').click()");
            assertEquals("true",evaluate(scenario,"!document.querySelector('#page-history').hidden"));
            scenario.recreate();waitFor(scenario,"window.Luma && state.ready && state.history.length>0 && document.documentElement.dataset.theme==='dark'");
            assertEquals("true",evaluate(scenario,"state.history[0].outputCount===2 && state.selection.length===0"));
            evaluate(scenario,"send('theme',{mode:'system'})");
        }
    }
}
