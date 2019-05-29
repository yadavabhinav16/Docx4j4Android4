package docx4j.org.docx4j4android4;

import android.Manifest;
import android.app.Activity;
import android.app.ActivityManager;
import android.content.ComponentName;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.os.Environment;
import android.support.design.widget.FloatingActionButton;
import android.support.design.widget.Snackbar;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.support.v7.widget.Toolbar;
import android.view.View;
import android.view.Menu;
import android.view.MenuItem;
import android.webkit.WebView;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;

import com.sun.xml.bind.v2.runtime.JAXBContextImpl;

import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.convert.out.html.AbstractHtmlExporter;
import org.docx4j.convert.out.html.HtmlExporterNG2;
import org.docx4j.convert.out.html.HtmlExporterNonXSLT;
import org.docx4j.jaxb.Context;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.model.fields.FieldUpdater;
import org.docx4j.model.images.ConversionImageHandler;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.io.LoadFromZipNG;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.CTCompat;
import org.docx4j.wml.Document;
import org.docx4j.wml.Text;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import java.io.InputStream;

import org.docx4j.XmlUtils;
import org.docx4j.convert.out.html.HtmlExporterNonXSLT;
import org.docx4j.model.images.ConversionImageHandler;
import org.docx4j.openpackaging.io.LoadFromZipNG;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;


public class MainActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        Toolbar toolbar = (Toolbar) findViewById(R.id.toolbar);
        setSupportActionBar(toolbar);

        FloatingActionButton fab = (FloatingActionButton) findViewById(R.id.fab);
        fab.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                Snackbar.make(view, "Replace with your own action", Snackbar.LENGTH_LONG)
                        .setAction("Action", null).show();
            }
        });

        try {
            System.out.println("about to create package");

            // org.apache.harmony.xml.parsers.SAXParserFactoryImpl throws SAXNotRecognizedException
            // for feature http://javax.xml.XMLConstants/feature/secure-processing
            // so either disable XML security, or use a different parser.  Here we disable it.
            org.docx4j.jaxb.ProviderProperties.getProviderProperties().put(JAXBContextImpl.DISABLE_XML_SECURITY, Boolean.TRUE);

            // Can we init the Context?
            // You can delete this if you want...
            System.out.println(Context.getJaxbImplementation());
            System.out.println(Context.jc.getClass().getName());

            // Create WordprocessingMLPackage
            WordprocessingMLPackage w = WordprocessingMLPackage.createPackage();
            w.getMainDocumentPart().setJaxbElement(new Document() );
            w.getMainDocumentPart().addParagraphOfText("hello from android");

            // Test marshalling works
            String XML = XmlUtils.marshaltoString(w.getMainDocumentPart().getJaxbElement());
            // or simply w.getMainDocumentPart().getXML();

            System.out.println(XML);

            // Test unmarshalling works
            try {
                Object o = XmlUtils.unmarshalString(XML, Context.jc);
                System.out.println(o.getClass().getName());
            } catch (JAXBException e) {
                e.printStackTrace();
            }

            // Test XPath works
            try {
                java.util.List<Object> results = w.getMainDocumentPart().getJAXBNodesViaXPath("//w:t", false);
                System.out.println("Xpath result count:" + results.size());
            } catch (JAXBException e) {
                e.printStackTrace();
            } catch (XPathBinderAssociationIsPartialException e) {
                e.printStackTrace();
            }

            System.out.println("done!");

        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////


        Button button1 = (Button) findViewById(R.id.button1);
        button1.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                // your handler code here
                WordprocessingMLPackage wordMLPackage = null;
                try {
                    wordMLPackage = WordprocessingMLPackage.createPackage();
                } catch (InvalidFormatException e) {
                    e.printStackTrace();
                }
                MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
                EditText edit = (EditText)findViewById(R.id.inputSave);
                String result = edit.getText().toString();

                mdp.addParagraphOfText(result);

                DocumentSettingsPart dsp = null;
                try {
                    dsp = mdp.getDocumentSettingsPart(true);
                } catch (InvalidFormatException e) {
                    e.printStackTrace();
                }
                CTCompat compat = Context.getWmlObjectFactory().createCTCompat();
                try {
                    dsp.getContents().setCompat(compat);
                } catch (Docx4JException e) {
                    e.printStackTrace();
                }
                compat.setCompatSetting("compatibilityMode", "http://schemas.microsoft.com/office/word", "15");




                // ActivityCompat.requestPermissions(MainActivity.this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, 55);

                //Activity activity = (Activity) Context;
                if (ContextCompat.checkSelfPermission(MainActivity.this,
                        Manifest.permission.WRITE_EXTERNAL_STORAGE)
                        != PackageManager.PERMISSION_GRANTED) {


                    // No explanation needed; request the permission
                    ActivityCompat.requestPermissions(MainActivity.this,
                            new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE},
                            56);

                    // MY_PERMISSIONS_REQUEST_READ_CONTACTS is an
                    // app-defined int constant. The callback method gets the
                    // result of the request.

                } else {}







                File path= Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS);
                //String filename = System.getProperty("user.dir") + "/OUT_hello_new.docx";
                //path+=+ "/OUT_hello_new.docx";
                File xmlFile = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS) + "/OUT_hello_new.docx");
                String showPath=xmlFile.getPath();
                System.out.println(showPath);
                String saveName=path.getPath()+ "/OUT_hello_new.docx";
                try {
                    Docx4J.save(wordMLPackage, xmlFile, Docx4J.FLAG_SAVE_ZIP_FILE);
                } catch (Docx4JException e) {
                    e.printStackTrace();
                }
                System.out.println("Saved " + saveName);
            }
        });




        ///// writing ends here




        /////////////////////////////////////////////////////////////////////////////////////////////////////

        ////button2 part of loading into textview

        Button button2 = (Button) findViewById(R.id.button2);
        button2.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {

                EditText edit = (EditText)findViewById(R.id.fileName);
                String result2 = edit.getText().toString();
                String finalText="";
                try {
                    File fl = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)+"/" + result2 +".docx");
                    WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(fl);
                    MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

                    String textNodesXPath = "//w:t";
                    List<Object> textNodes= documentPart
                            .getJAXBNodesViaXPath(textNodesXPath, true);
                    for (Object obj : textNodes) {
                        Text text = (Text) ((JAXBElement) obj).getValue();
                        String textValue = text.getValue();
                        finalText+=textValue;


                    }
                    TextView tv= (TextView)findViewById(R.id.displayText);
                    tv.setText(finalText);




                } catch (Docx4JException e1) {
                    e1.printStackTrace();
                } catch (JAXBException e1) {
                    e1.printStackTrace();
                }
            }});







        /////////////////////////////////////////////////////////////////////////////////////////////////////



    }

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        // Inflate the menu; this adds items to the action bar if it is present.
        getMenuInflater().inflate(R.menu.menu_main, menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        // Handle action bar item clicks here. The action bar will
        // automatically handle clicks on the Home/Up button, so long
        // as you specify a parent activity in AndroidManifest.xml.
        int id = item.getItemId();

        //noinspection SimplifiableIfStatement
        if (id == R.id.action_settings) {
            return true;
        }

        return super.onOptionsItemSelected(item);
    }
}
