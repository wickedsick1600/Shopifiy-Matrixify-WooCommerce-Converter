cd path/to/glossylounge-migration
npm install
node convert.js path/to/your-matrixify-export.xlsx woocommerce-import.csv
# Dry-run (no file written):
node convert.js path/to/export.xlsx out.csv --dry-run