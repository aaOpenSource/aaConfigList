# aaConfigList

Exports a Galaxy's configuration information using SQL queries. 
See the example file from the stock Reactor Demo Galaxy: [Demo_Area.txt](Log/Demo_Area.txt).

The code is the first phase of a version control. When run it will generate files of the configuration of
all instances in the areas under the base area being requested. Each file represents an area. At this point it
is capable of gathering a listing of all Areas with version, Object instances with version, UDA's with extension flags,
inherited scripts. There needs to be a file structure set up for it to operate in and all of the files attached need
to have the .vbs extension added to them. This can be run from any system that is capable of accessing the GR node.
The file structure should be placed in the root, or as close as possible, of whichever drive you choose to locate it
on. If the galaxy has too many levels, and the names get too long I have run into path\file name length limitations. 

## Usage

Place the Visual Basic script files in a root directory on a machine that can
access the GR node. These scripts use a direct SQL connection, so MX or GRAccess is not required.

You will need the SQL database, user name, password, and root area that is going to be exported.

## Platforms

* Tested on System Platform 3.0 SP2, and 2014.

## Runtime Notes

Running aaConfigList with the NoAttrib argument will result in an abbreviated listing with objects with versions only. 
This runs much, much faster and can identify which objects have changed when compared with an earlier NoAttrib run. 

Double click the aaConfigList.vbs. A message box should appear asking for the name of the Server(GRNode). Next it will
ask the name of the database(Galaxy). Then the name of the User. This user must have read access to the DB.It will then
need the SQLServer password for the user. Then it will ask for the base area. Use the hierarchical name of the highest
level area you want documented. Next it will ask for optional arguments. Use the //x if you would like to invoke VS.

## Contributors

* The code was emailed to me (Eliot Landrum) on an anonymous basis.