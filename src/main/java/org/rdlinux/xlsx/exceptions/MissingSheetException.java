package org.rdlinux.xlsx.exceptions;

public class MissingSheetException extends RuntimeException {

    private static final long serialVersionUID = -9001714880458246425L;

    public MissingSheetException() {
        super();
    }

    public MissingSheetException(String msg) {
        super(msg);
    }

    public MissingSheetException(Exception e) {
        super(e);
    }

    public MissingSheetException(String msg, Exception e) {
        super(msg, e);
    }
}
