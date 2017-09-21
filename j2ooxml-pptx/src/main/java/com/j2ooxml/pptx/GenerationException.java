package com.j2ooxml.pptx;

public class GenerationException extends Exception {

    private static final long serialVersionUID = -4921752963711371513L;

    public GenerationException(String message) {
        super(message);
    }

    public GenerationException(Throwable cause) {
        super(cause);
    }

    public GenerationException(String message, Throwable cause) {
        super(message, cause);
    }
}
