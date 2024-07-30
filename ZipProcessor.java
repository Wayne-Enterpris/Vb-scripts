package com.example;

import org.apache.log4j.Logger;

import java.io.*;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class ZipProcessor {
    private static final Logger logger = Logger.getLogger(ZipProcessor.class);
    private static final String FILE_PATTERN = ".*\\.txt"; // Adjust the file pattern as needed

    public static void main(String[] args) {
        if (args.length != 1) {
            logger.error("Usage: java -jar zip-processor.jar <path-to-zipfilelist.txt>");
            System.exit(1);
        }

        String zipFileListPath = args[0];

        try (BufferedReader br = new BufferedReader(new FileReader(zipFileListPath))) {
            String zipFileName;
            while ((zipFileName = br.readLine()) != null) {
                File zipFile = new File(zipFileName);
                if (!zipFile.exists() || !zipFile.isFile()) {
                    logger.warn("File not found: " + zipFileName);
                    continue;
                }

                processZipFile(zipFile);
            }
        } catch (IOException e) {
            logger.error("Error reading zip file list", e);
        }
    }

    private static void processZipFile(File zipFile) {
        int fileCount = 0;

        try (ZipFile zf = new ZipFile(zipFile)) {
            Enumeration<? extends ZipEntry> entries = zf.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                if (!entry.isDirectory() && entry.getName().matches(FILE_PATTERN)) {
                    File extractedFile = new File("extracted_" + zipFile.getName() + "_" + entry.getName());
                    try (InputStream is = zf.getInputStream(entry);
                         OutputStream os = new FileOutputStream(extractedFile)) {

                        byte[] buffer = new byte[1024];
                        int length;
                        while ((length = is.read(buffer)) > 0) {
                            os.write(buffer, 0, length);
                        }

                        logger.info("Extracted: " + extractedFile.getAbsolutePath());
                        fileCount++;
                    }
                }
            }
        } catch (IOException e) {
            logger.error("Error processing zip file: " + zipFile.getName(), e);
        }

        logger.info("Total files extracted from " + zipFile.getName() + ": " + fileCount);
    }
}
