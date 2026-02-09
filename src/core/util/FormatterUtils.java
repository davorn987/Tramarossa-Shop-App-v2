package core.util;

import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;

public class FormatterUtils {
    public static DecimalFormat getItalianDecimalFormat() {
        DecimalFormatSymbols symbols = new DecimalFormatSymbols();
        symbols.setDecimalSeparator(',');
        return new DecimalFormat("#,##0.00", symbols);
    }
}
