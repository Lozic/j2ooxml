package com.j2ooxml.pptx;

import java.io.IOException;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.w3c.css.sac.InputSource;
import org.w3c.dom.Document;
import org.w3c.dom.css.CSSStyleSheet;

import com.j2ooxml.pptx.util.XmlUtil;
import com.steadystate.css.parser.CSSOMParser;

public class PptxGenerator {

    private VariableProcessor variableProcessor = new VariableProcessor();

    private ImageReplacer imageReplacer = new ImageReplacer();

    private VideoReplacer videoReplacer = new VideoReplacer();

    private Deletor deletor = new Deletor();

    private EmptyNodesRemover emptyNodesRemover = new EmptyNodesRemover();

    public void process(Path templatePath, Path cssPath, Path outputPath, Map<String, Object> model)
            throws IOException, GenerationException {
        try {
            Files.copy(templatePath, outputPath, StandardCopyOption.REPLACE_EXISTING);
            FileSystem fs = FileSystems.newFileSystem(outputPath, null);

            Path slides = fs.getPath("/ppt/slides");
            try (DirectoryStream<Path> directoryStream = Files.newDirectoryStream(slides)) {
                for (Path slideXml : directoryStream) {
                    if (Files.isRegularFile(slideXml)) {
                        String slide = slideXml.getFileName().toString();
                        Path relXml = fs.getPath("/ppt/slides/_rels/" + slide + ".rels");
                        CSSOMParser parser = new CSSOMParser();
                        String cssString = new String(Files.readAllBytes(cssPath), StandardCharsets.UTF_8);
                        if (model.containsKey("CSS")) {
                            String modelCss = (String) model.get("CSS");
                            cssString += " " + modelCss;
                        }
                        StringReader reader = new StringReader(cssString);
                        CSSStyleSheet css = parser.parseStyleSheet(new InputSource(reader), null, null);
                        State state = imageReplacer.replace(fs, slideXml, model, css);
                        videoReplacer.replace(fs, slideXml, state, model);
                        deletor.process(state, model);
                        variableProcessor.process(state, css, model);
                        XmlUtil.save(slideXml, state.getSlideDoc());
                        XmlUtil.save(relXml, state.getRelDoc());
                        try (Stream<String> lines = Files.lines(relXml, StandardCharsets.UTF_8)) {
                            List<String> replaced = lines
                                    .map(line -> line.replaceAll("(Target=\".*?\"(\\sTargetMode=\".*?\")?)\\s(Type=\".*?\")", "$3 $1"))
                                    .collect(Collectors.toList());
                            lines.close();
                            Files.delete(relXml);
                            Files.write(relXml, replaced, StandardCharsets.UTF_8);
                        }
                    }
                }
            }

            slides = fs.getPath("/ppt/slideLayouts");
            try (DirectoryStream<Path> directoryStream = Files.newDirectoryStream(slides)) {
                for (Path slideXml : directoryStream) {
                    if (Files.isRegularFile(slideXml)) {
                        Document slideDoc = XmlUtil.parse(slideXml);
                        emptyNodesRemover.process(slideDoc, model);
                        XmlUtil.save(slideXml, slideDoc);
                    }
                }
            }
            Properties.fillProperies(fs, model);
            fs.close();
        } catch (Exception e) {
            throw new GenerationException("Could not generate resulting ppt.", e);
        }
    }

}
