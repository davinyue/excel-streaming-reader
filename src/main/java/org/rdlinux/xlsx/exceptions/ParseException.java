package org.rdlinux.xlsx.exceptions;

public class ParseException extends RuntimeException {

    private static final long serialVersionUID = 5517822224954078559L;

    public ParseException() {
        super();
    }

    public ParseException(String msg) {
        super(msg);
    }

    public ParseException(Exception e) {
        super(e);
    }

    public ParseException(String msg, Exception e) {
        super(msg, e);
    }
}
