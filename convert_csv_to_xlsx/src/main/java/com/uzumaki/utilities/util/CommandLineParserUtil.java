package com.uzumaki.utilities.util;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class CommandLineParserUtil {

	/**
     * Equals delimiter.
     */
    private static final String EQUALS_DELIMITER = "=";

    /**
     * Comma delimiter.
     */
    private static final String COMMA_DELIMITER = ",";

    /**
     * Quote delimiter.
     */
    private static final String QUOTE_DELIMITER = "\"";

    /**
     * Constructor.
     */
    private CommandLineParserUtil()
    {
    }

    /**
     * Parses the command line argument of the form "key=value".
     * The method splits the argument on the "=" sign. It then returns the "key" from the argument.
     * @param argument: Command line argument of the form key=value.
     * @return The key in the argument.
     */
    public static String getCommandLineArgKey( String argument )
    {
        if ( argument == null )
        {
            return null;
        }

        String key = null;
        String[] splitArgs = argument.split( EQUALS_DELIMITER );
        if ( splitArgs.length == 2 )
        {
            key = splitArgs[0];
        }
        return key;
    }

    /**
     * Parses the command line argument of the form "key=value".
     * The method splits the argument on the "=" sign. It then returns the "value" from the argument.
     * @param argument: Command line argument of the form key=value.
     * @return The value in the argument.
     */
    public static String getCommandLineArgValue( final String argument )
    {
        if ( argument == null )
        {
            return null;
        }
        String value = null;
        String[] splitArgs = argument.split( EQUALS_DELIMITER );
        if ( splitArgs.length == 2 )
        {
            value = splitArgs[1];
            //if the value is in quotes, then strip the quotes out.
            int q0 = value.indexOf( QUOTE_DELIMITER );
            int q1 = value.lastIndexOf( QUOTE_DELIMITER );
            if ( 0 <= q0  && 0 <= q1 && q0 < q1 )
            {
                value = value.substring( q0 + 1, q1 );
                value = value.trim();
            }
        }
        return value;
    }

    /**
     * Gets the value for a given key from the command line arguments.
     * @param args: The command line arguments.
     * @param key: The key for which the value has to be fetched.
     * @return value, can be null;
     */
    public static String getValueForKey( String[] args, String key )
    {
        if ( args == null || key == null )
        {
            return null;
        }
        String value = null;
        for ( int i = 0; i < args.length; i++ )
        {
            String tmpKey = getCommandLineArgKey( args[i].trim() );
            if ( tmpKey != null && tmpKey.equals( key ) )
            {
                value = CommandLineParserUtil.getCommandLineArgValue( args[i].trim() );
                break;
            }
        }
        return value;
    }

    /**
     * Checks whether the command line argument have help option.
     * @param args: Command line arguments.
     * @return true if -h option is in arguments list, false otherwise.
     */
    public static boolean isHelpOption( String[] args )
    {
        if ( args == null )
        {
            return true;
        }
        return isOptionExists( args, "-h" ); //$NON-NLS-1$
    }

    /**
     * Checks whether the command line arguments have the given option.
     * @param args: The command line arguments.
     * @param option: The option to be verified
     * @return true if the option exists in the command line arguments. Otherwise returns false
     */
    public static boolean isOptionExists( String[] args, String option )
    {
        boolean exists = false;
        if ( args == null )
        {
            return exists;
        }

        for ( int i = 0; i < args.length; i++ )
        {
            if ( args[i].trim().equals( option ) )
            {
                exists = true;
                break;
            }
        }
        return exists;
    }

    /**
     * Splits a comma separated string.
     * @param commaSeparatedString: Comma separated string.
     * @return returns the key in the argument.
     */
    public static String[] splitStringOnComma( final String commaSeparatedString )
    {
        if ( commaSeparatedString == null || commaSeparatedString.length() <= 0 )
        {
            return null;
        }

        return commaSeparatedString.split( COMMA_DELIMITER );
    }

    /**
     * Gets the comma separated strings from the value for a specified key and builds a list of strings.
     * @param args: The command line arguments.
     * @param key: The key in the argument list that contains a list comma separated strings
     * @return List of String
     */
    public static List<String> buildList( String[] args, String key )
    {
        List<String> list = null;
        String stringList = CommandLineParserUtil.getValueForKey( args, key );
        if ( stringList != null )
        {
            list = new ArrayList<String>();
            String[] splitArgs = CommandLineParserUtil.splitStringOnComma( stringList );
            for ( int k = 0; k < splitArgs.length; k++ )
            {
                list.add( splitArgs[k].trim() );
            }
        }
        return list;
    }
    
    /**
     * This method validates in the arguments provided in args are the supported set of arguments.
     * @param supportedArguments: The arguments supported by utility
     * @param args: The args to validate
     * @return the list containing invalid arguments or null if all are valid 
     */
    public static List<String> getInvalidArguments( String [] supportedArguments, String[] args )
    {
        List<String> invalidArglist = new ArrayList<String>();
        Arrays.sort( supportedArguments );
        
        for( int i=0; i < args.length ; i++ )
        {
            String arg = args[i];
            String [] temp = arg.split( EQUALS_DELIMITER );
            if( temp != null && temp.length > 0 )
            {
                if( Arrays.binarySearch( supportedArguments, temp[0] ) < 0 )
                {
                    invalidArglist.add( temp[0] );
                }
            }
        }
        if( invalidArglist.size() == 0 )
        {
            return null;
        }
        else
        {
            return invalidArglist;    
        }
    }
}
