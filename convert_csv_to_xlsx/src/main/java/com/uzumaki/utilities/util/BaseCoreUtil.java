package com.uzumaki.utilities.util;

import java.io.File;

public class BaseCoreUtil
{
    /**
     * Deletes all files and sub-directories under the given directory.
     * @param dir: The directory to be deleted
     * @return true, if deletion is successful; false, otherwise
     */
    public static boolean deleteDir( File dir )
    {
        if ( dir.isDirectory() )
        {
            // delete all files and sub-directories first
            String[] children = dir.list();
            for ( int i = 0; i < children.length; i++ )
            {
                boolean success = deleteDir( new File( dir, children[i] ) );
                if ( !success )
                {
                    return false;
                }
            }
        }

        // delete the empty directory or file
        return dir.delete();
    }
}
