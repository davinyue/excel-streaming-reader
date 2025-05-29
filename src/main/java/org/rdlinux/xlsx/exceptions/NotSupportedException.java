package org.rdlinux.xlsx.exceptions;

public class NotSupportedException extends RuntimeException {

    private static final long serialVersionUID = 3698767908427144923L;

    public NotSupportedException() {
        super();
    }

    public NotSupportedException(String msg) {
        super(msg);
    }

    public NotSupportedException(Exception e) {
        super(e);
    }

    public NotSupportedException(String msg, Exception e) {
        super(msg, e);
    }
}
