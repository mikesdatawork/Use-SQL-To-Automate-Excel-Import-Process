![MIKES DATA WORK GIT REPO](https://raw.githubusercontent.com/mikesdatawork/images/master/git_mikes_data_work_banner_01.png "Mikes Data Work")        

# Use SQL To Automate Excel Import Process
**Post Date: February 20, 2018** 
![SQL Import](https://mikesdatawork.files.wordpress.com/2018/02/xls_screen_01.png "SQL Excel Import")   


## Contents    
- [About Process](##About-Process)  
- [SQL Logic](#SQL-Logic)  
- [Build Info](#Build-Info)  
- [Author](#Author)  
- [License](#License)       

## About-Process

<p>Here's an example of something you can write to import Excel Data directly into SQL Server without the use of any other data services.</p>    


## SQL-Logic
```SQL
use [master];
set nocount on
 
-- allow queries to directly access excel files
exec [master]..sp_configure 'show advanced options', 1;
reconfigure with override;
exec [master]..sp_configure 'ad hoc distributed queries', 1;
reconfigure with override;
exec [master].dbo.sp_msset_oledb_prop N'microsoft.ace.oledb.12.0', N'allowinprocess', 1;
exec [master].dbo.sp_msset_oledb_prop N'microsoft.ace.oledb.12.0', N'dynamicparameters', 1;
 
use [compliance];
set nocount on
 
-- get simple file list
declare @import_path    varchar(255) = 'C:\SQLIMPORTS\'
declare @files      table ([subdirectory] varchar(255), [depth] int, [file] int)
insert into @files exec master..xp_dirtree @import_path, 1, 1; delete from @files where [file] = 0 
  
if object_id('tempdb..#import_folder') is not null
drop table      #import_folder
create table    #import_folder  ([file_info] nvarchar(255))
declare         @file_list  table ([file_name] nvarchar(255))
  
-- get file meta data
Insert into     #import_folder exec xp_cmdshell 'dir C:\SQLIMPORTS\*.*'
delete from     #import_folder where [file_info] is null 
delete from     #import_folder where [file_info] like '%dir%' 
delete from     #import_folder where [file_info] like '%volume%' 
delete from     #import_folder where [file_info] like '%bytes%'
  
-- combine meta data with file name
if object_id('tempdb..#files_found') is not null
drop table      #files_found
create table    #files_found ([create_date] datetime, [file_name] nvarchar(255))
 
-- compile list of all files found.
declare @compile_list       varchar(max)
set     @compile_list       = ''
select  @compile_list       = @compile_list + 
'insert into #files_found  ([create_date], [file_name]) select cast(left([file_info], 20) as datetime), ''' + [subdirectory] + ''' from #import_folder where [file_info] like ''%' + [subdirectory] + '%'';' + char(10)
from    @files
exec    (@compile_list)
  
 -- create table name based on latest file added to directory
declare @new_file   varchar(255) = (select top 1 [file_name] from #files_found where [file_name] like '%.xlsx' and left([file_name], 8) like 'MSSQL FY%' order by [create_date] desc)
declare @new_table  varchar(255) = (select 'IMPORT_' + replace(substring(@new_file, 7, charindex('Q', @new_file) + 3), ' ', '') + '_00')
 
-- create next import table based on excel columns found.
declare @target varchar(255) = (select top 1 [table_name] from information_schema.tables where [table_name] like 'IMPORT_FY%' order by [table_name] desc)
declare @next_table varchar(255) = (select isnull(@target, @new_table))
select  @next_table = upper(left(@new_table, 14)) + format(cast(right(@next_table, 2) as int) + 1, '00')
 
declare @create_table   varchar(max) = 'create table [' + @next_table + ']' + char(10) + 
'
(
   [F1]  nvarchar(max),  [F2]  nvarchar(max),  [F3] nvarchar(max),   [F4]  nvarchar(max)
,  [F5]  nvarchar(max),  [F6]  nvarchar(max),  [F7] nvarchar(max),   [F8]  nvarchar(max)
,  [F9]  nvarchar(max),  [F10] nvarchar(max),  [F11] nvarchar(max),  [F12] nvarchar(max)
,  [F13] nvarchar(max),  [F14] nvarchar(max),  [F15] nvarchar(max),  [F16] nvarchar(max)
,  [F17] nvarchar(max),  [F18] nvarchar(max),  [F19] nvarchar(max),  [F20] nvarchar(max)
,  [F21] nvarchar(max),  [F22] nvarchar(max),  [F23] nvarchar(max),  [F24] nvarchar(max)
,  [F25] nvarchar(max),  [F26] nvarchar(max),  [F27] nvarchar(max),  [F28] nvarchar(max)
,  [F29] nvarchar(max),  [F30] nvarchar(max),  [F31] nvarchar(max),  [F32] nvarchar(max)
,  [F33] nvarchar(max),  [F34] nvarchar(max),  [F35] nvarchar(max),  [F36] nvarchar(max)
,  [F37] nvarchar(max),  [F38] nvarchar(max),  [F39] nvarchar(max),  [F40] nvarchar(max)
,  [F41] nvarchar(max),  [F42] nvarchar(max),  [F43] nvarchar(max),  [F44] nvarchar(max)
,  [F45] nvarchar(max),  [F46] nvarchar(max),  [F47] nvarchar(max),  [F48] nvarchar(max)
,  [F49] nvarchar(max),  [F50] nvarchar(max),  [F51] nvarchar(max),  [F52] nvarchar(max)
,  [F53] nvarchar(max),  [F54] nvarchar(max),  [F55] nvarchar(max),  [F56] nvarchar(max)
,  [F57] nvarchar(max),  [F58] nvarchar(max),  [F59] nvarchar(max),  [F60] nvarchar(max)
,  [F61] nvarchar(max),  [F62] nvarchar(max),  [F63] nvarchar(max),  [F64] nvarchar(max)
,  [F65] nvarchar(max),  [F66] nvarchar(max),  [F67] nvarchar(max),  [F68] nvarchar(max)
,  [F69] nvarchar(max),  [F70] nvarchar(max),  [F71] nvarchar(max),  [F72] nvarchar(max)
,  [F73] nvarchar(max),  [F74] nvarchar(max),  [F75] nvarchar(max),  [F76] nvarchar(max)
,  [F77] nvarchar(max),  [F78] nvarchar(max),  [F79] nvarchar(max),  [F80] nvarchar(max)
,  [F81] nvarchar(max),  [F82] nvarchar(max),  [F83] nvarchar(max),  [F84] nvarchar(max)
,  [F85] nvarchar(max),  [F86] nvarchar(max),  [F87] nvarchar(max),  [F88] nvarchar(max)
,  [F89] nvarchar(max),  [F90] nvarchar(max),  [F91] nvarchar(max),  [F92] nvarchar(max)
,  [F93] nvarchar(max),  [F94] nvarchar(max),  [F95] nvarchar(max),  [F96] nvarchar(max)
,  [F97] nvarchar(max),  [F98] nvarchar(max),  [F99] nvarchar(max),  [F100] nvarchar(max)
,  [F101] nvarchar(max), [F102] nvarchar(max), [F103] nvarchar(max)
)'
exec (@create_table)
 
-- populate table with valuse from excel sheet.
declare @populate_table varchar(max) = ('insert into [' + @next_table + '] select * from openrowset(''Microsoft.ACE.OLEDB.12.0'', ''Excel 12.0 Xml; HDR=YES ;Database=' + @import_path + @new_file + ''', ''SELECT * FROM [MSSQL$]'')')
exec (@populate_table)
 
-- create final table after import table is created.
declare @final_table    varchar(255) = (select right(@next_table,  9))
declare @final_build    varchar(max) = ('create table [' + @final_table + ']' + char(10) + 
'(
    [Issue]             nvarchar(255) --F1
,   [Severity]          nvarchar(255) --F19
,   [Control_Mapping]       nvarchar(555) --F12
,   [Condition]         nvarchar(max) --F3
,   [Server]            nvarchar(255) --F6
,   [Module]            nvarchar(255) --F5
,   [Version]           nvarchar(255) --F8
,   [Cause]             nvarchar(max) --F9
,   [Recommendation]        nvarchar(max) --F10, F11
,   [Comments]          nvarchar(max) --F11
)')
exec    (@final_build)
 
-- populate final table.   map F# columns from import table to final table and perfrom insert process to final table.
declare @populate_final varchar(max) = ('insert into [' + @final_table + '] 
select [F1], [F19], [F12], [F3], upper([F6]), replace(upper([F5]), ''.MyDomain.com'', ''''), [F8], [F9], ([F10] + ''  '' +  [F11]), [F11]
from [' + @next_table + '] where [F1] not in (''Issue #'') and [F1] is not null order by [F1] asc; update [' + @final_table + '] set [comments] = NULL;')
exec    (@populate_final)
```




[![WorksEveryTime](https://forthebadge.com/images/badges/60-percent-of-the-time-works-every-time.svg)](https://shitday.de/)

## Build-Info

| Build Quality | Build History |
|--|--|
|<table><tr><td>[![Build-Status](https://ci.appveyor.com/api/projects/status/pjxh5g91jpbh7t84?svg?style=flat-square)](#)</td></tr><tr><td>[![Coverage](https://coveralls.io/repos/github/tygerbytes/ResourceFitness/badge.svg?style=flat-square)](#)</td></tr><tr><td>[![Nuget](https://img.shields.io/nuget/v/TW.Resfit.Core.svg?style=flat-square)](#)</td></tr></table>|<table><tr><td>[![Build history](https://buildstats.info/appveyor/chart/tygerbytes/resourcefitness)](#)</td></tr></table>|

## Author

[![Gist](https://img.shields.io/badge/Gist-MikesDataWork-<COLOR>.svg)](https://gist.github.com/mikesdatawork)
[![Twitter](https://img.shields.io/badge/Twitter-MikesDataWork-<COLOR>.svg)](https://twitter.com/mikesdatawork)
[![Wordpress](https://img.shields.io/badge/Wordpress-MikesDataWork-<COLOR>.svg)](https://mikesdatawork.wordpress.com/)

    
## License
[![LicenseCCSA](https://img.shields.io/badge/License-CreativeCommonsSA-<COLOR>.svg)](https://creativecommons.org/share-your-work/licensing-types-examples/)

![Mikes Data Work](https://raw.githubusercontent.com/mikesdatawork/images/master/git_mikes_data_work_banner_02.png "Mikes Data Work")

