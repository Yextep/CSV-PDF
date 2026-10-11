package com.luma.app;

import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Paint;
import android.graphics.Typeface;
import android.graphics.pdf.PdfDocument;
import android.text.Layout;
import android.text.StaticLayout;
import android.text.TextPaint;
import org.json.JSONArray;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/** PDF A4 de texto, con fuentes Android, saltos de página y caracteres Unicode. */
final class PdfExporter {
    static void write(JSONArray lines, String title, File target) throws Exception {
        StringBuilder text = new StringBuilder();
        for (int i = 0; i < lines.length(); i++) {
            if (i > 0) text.append('\n');
            text.append(lines.getString(i).replace("\t", "    "));
        }
        TextPaint body = new TextPaint(Paint.ANTI_ALIAS_FLAG);
        body.setColor(Color.rgb(35, 32, 53));
        body.setTextSize(11);
        body.setTypeface(Typeface.create("sans-serif", Typeface.NORMAL));
        StaticLayout layout = StaticLayout.Builder.obtain(text, 0, text.length(), body, 499)
                .setAlignment(Layout.Alignment.ALIGN_NORMAL).setIncludePad(false)
                .setLineSpacing(4, 1).build();
        List<int[]> pages = new ArrayList<>();
        int line = 0;
        while (line < layout.getLineCount()) {
            int start = line, top = layout.getLineTop(start);
            while (line < layout.getLineCount() && layout.getLineBottom(line) - top <= 710) line++;
            if (line == start) line++;
            pages.add(new int[]{top, layout.getLineTop(line) - top});
            if (pages.size() > 300) throw new IOException("El documento supera 300 páginas de texto. Divide el archivo.");
        }
        File temp = new File(target.getParentFile(), target.getName() + ".part");
        PdfDocument pdf = new PdfDocument();
        try {
            Paint heading = new Paint(Paint.ANTI_ALIAS_FLAG);
            heading.setColor(Color.rgb(113, 94, 150));
            heading.setTextSize(9);
            for (int i = 0; i < pages.size(); i++) {
                PdfDocument.Page page = pdf.startPage(new PdfDocument.PageInfo.Builder(595, 842, i + 1).create());
                Canvas canvas = page.getCanvas();
                String headingText = title;
                while (headingText.length() > 1 && heading.measureText(headingText) > 440) {
                    headingText = headingText.substring(0, headingText.length() - 2) + "…";
                }
                canvas.drawText(headingText, 48, 43, heading);
                canvas.drawLine(48, 54, 547, 54, heading);
                canvas.save();
                canvas.clipRect(48, 72, 547, 72 + pages.get(i)[1]);
                canvas.translate(48, 72 - pages.get(i)[0]);
                layout.draw(canvas);
                canvas.restore();
                canvas.drawText((i + 1) + " / " + pages.size(), 510, 809, heading);
                pdf.finishPage(page);
            }
            try (FileOutputStream output = new FileOutputStream(temp)) {
                pdf.writeTo(output);
                output.getFD().sync();
            }
            if (!temp.renameTo(target)) throw new IOException("No se pudo guardar el PDF.");
        } finally {
            pdf.close();
            if (temp.exists()) temp.delete();
        }
    }
}
