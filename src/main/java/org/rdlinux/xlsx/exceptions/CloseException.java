package org.rdlinux.xlsx.exceptions;

public class CloseException extends RuntimeException {

    private static final long serialVersionUID = 4243405987879871932L;

    public CloseException() {
        super();
    }

    public CloseException(String msg) {
        super(msg);
    }

    public CloseException(Exception e) {
        super(e);
    }

    public CloseException(String msg, Exception e) {
        super(msg, e);
    }
}
