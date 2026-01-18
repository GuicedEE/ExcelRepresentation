package com.guicedee.services.excelrepresentation;


import java.io.Serial;

/**
 * Runtime exception used to signal failures while reading or rendering Excel
 * spreadsheets, including parse errors, stream issues, or unsupported formats.
 */
public class ExcelRenderingException extends RuntimeException {
    @Serial
    private static final long serialVersionUID = 1L;

    /**
     * Creates an exception with no detail message.
     */
    public ExcelRenderingException()
    {
    }

    /**
     * Creates an exception with a detail message.
     *
     * @param message
     * 		the detail message
     */
    public ExcelRenderingException(String message)
    {
        super(message);
    }

    /**
     * Creates an exception with a detail message and cause.
     *
     * @param message
     * 		the detail message
     * @param cause
     * 		the underlying cause
     */
    public ExcelRenderingException(String message, Throwable cause)
    {
        super(message, cause);
    }

    /**
     * Creates an exception with a cause.
     *
     * @param cause
     * 		the underlying cause
     */
    public ExcelRenderingException(Throwable cause)
    {
        super(cause);
    }

    /**
     * Creates an exception with full control over suppression and stack trace
     * writability.
     *
     * @param message
     * 		the detail message
     * @param cause
     * 		the underlying cause
     * @param enableSuppression
     * 		whether suppression is enabled or disabled
     * @param writableStackTrace
     * 		whether the stack trace should be writable
     */
    public ExcelRenderingException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace)
    {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
