ms-officer

run mvn package to updates binary on runner

/runner/input is the files that will become data inputs that can be used for templates files, ex: {{#mpp-file-name.taskName}}
/runner/template is the files that will become template for generating output files, ex: reading data from mpp, to create xls file
/runner/output is the result files
/runner/bin is the library files for processing the whole thing
/runner/RUN.bat is executable script to be run by Windows users to start the application
/runner/RUN.sh is executable script to be run by Mac/Linux users to start the application

/runner will become the binary package published to users
