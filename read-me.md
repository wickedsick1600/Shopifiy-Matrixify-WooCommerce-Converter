HOW TO USE:
- Add collection-export.xlsx(all collection details from matrixify) and input.xlsx(all products export from matrixify) in the root folder
- cd path/to/glossylounge-migration
- npm install
- node convert.js
- Import products in woocommerce
- Run assign-size-variations.php in Codesnippets(woocommerce plugin)
- Run fix-product-slug-nphp in Codesnippets as well


# Dry-run (no file written):
node convert.js path/to/export.xlsx out.csv --dry-run