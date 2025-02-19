# vcard-cleaner

This .net console application:
 * loads either a vCard (.vcf) file or a csv file from the supplied path, 
 * parses each contact into objects, 
 * cleans the phone numbers if necessary, 
 * merges the records that share the same name but have different phone numbers, 
 * removes duplicates, 
 * writes the distinct records back to a new vcf file and/or a new csv file.
