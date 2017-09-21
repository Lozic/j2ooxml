package com.j2ooxml.pptx;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.commons.lang3.tuple.Pair;
import org.junit.Test;

public class PptGeneratorTest {

    private PptxGenerator generator = new PptxGenerator();

    @Test
    public void test() {
        try {
            Path templatePath = Paths.get(this.getClass().getResource("test.pptx").toURI());
            Path cssPath = Paths.get(this.getClass().getResource("test.css").toURI());
            Path outputPath = Paths.get("target/out.pptx");
            Map<String, Object> model = new HashMap<>();

            model.put("no-background", true);

            model.put("kf_title", null);
            model.put("kf_number1", "<span style='text-align: center;'>11</span>");
            model.put("asset_type", "Test sdfsdf sdfsdfsdfsd sdfsdf sdfsdfsdfsdf sdfsdfsddf sdfsdf");
            model.put("sol_name2", "<span style='text-align: right;'>name</span>");
            model.put("benefit_description", "X<sup>2</sup> H<sub>2</sub>O");

            model.put("sol_image1", Paths.get(this.getClass().getResource("1.png").toURI()));
            model.put("sol_image2", Paths.get(this.getClass().getResource("2.png").toURI()));
            model.put("sol_image3", Paths.get(this.getClass().getResource("3.png").toURI()));
            model.put("sol_image4", Paths.get(this.getClass().getResource("1.png").toURI()));
            model.put("template_image", Paths.get(this.getClass().getResource("mi.jpg").toURI()));

            model.put("companyName",
                    "E<b>p</b>a<i>m</i> <u>sy<b>st</b>ems</u>. <br /><span style='color: rgb(255, 0, 0);'>I'm red</span>");

            Slide slide = new Slide();
            slide.setTitle("no bullet<ul><li>first bullet</li><li>second bullet</li>"
                    + "<li>OMG, I'm <b>bold <u>underlined</u> <i>italic</i> bullet</b></li>"
                    + "<li>I'm half <b>bold</b> bullet</li></ul>again no bullet with <i>italic</i> end<br />"
                    + "<span style='font-size: 60.0pt;'>I'ma big,</span> <span style='font-size: 9pt;'>me-tiny,</span> I am standard<br/>"
                    + "<span class='test-class'>I'm styled in css by classname.</span>"
                    + "<ul class='cust'><li>Very custom</li><li>list very long text to test margin-left css property</li></ul>"
                    + "<a href='http://www.google.com'>Search on Google</a>");
            model.put("slide", slide);

            model.put("learnmore_description", "<ul class='triangle-bullets'>" + "<li><a>Visit us online</a></li>"
                    + "<li><a>Benchmark your performance</a></li>" + "<li><a>SAP Solution Explorer</a> </li></ul>");

            model.put("logo", Paths.get(this.getClass().getResource("logo.jpg").toURI()));
            model.put("custom", Paths.get(this.getClass().getResource("cust.jpg").toURI()));
            model.put("logoUrl", Paths.get(this.getClass().getResource("1.png").toURI()));

            Path video = Paths.get(this.getClass().getResource("test.mp4").toURI());
            Path thumb = Paths.get(this.getClass().getResource("1.png").toURI());
            Pair<Path, Path> videoPair = new ImmutablePair<>(thumb, video);
            model.put("video", videoPair);

            model.put("title", "T<sup>®</sup>T");

            model.put("CSS", "#logo {left: 100mm;}");

            model.put("SAPLogo", "left: 100mm;");

            model.put(Properties.CREATOR, "telepuzic");
            model.put(Properties.CREATED, LocalDateTime.now(ZoneOffset.UTC));

            generator.process(templatePath, cssPath, outputPath, model);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static class Slide {

        private String title;

        public String getTitle() {
            return title;
        }

        public void setTitle(String title) {
            this.title = title;
        }

    }
}
