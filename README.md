# SQLiteClass #

### AHK SQLite API Wrapper ###

AHK class providing support to access SQLite databases.


### Additional stuff ###

- [SQLite documentation](http://www.sqlite.org/docs.html)
- [SQLite download page](http://www.sqlite.org/download.html)

### Basic usage ###

- Create a new instance of the SQLiteClass calling `myDatabase := SQLiteClass()`

- Open your database calling `myDatabase.OpenDB(myDatabaseFilePath)`. If the file doesn't exist, a new database will be created unless you specify "false" as the third parameter.

- `myDatabase` object provides four methods to pass SQL statements to the database:
  - `myDatabase.Exec(sql)`
	Should be called for all SQL statements which don't return values from the database (e.g. CREATE, INSERT, UPDATE, etc.).

  - `myDatabase.GetTable(sql, Table, ...)`
	Should be called for SELECT statements whenever you want to get the complete result of the query as a "Table" object for direct access via the row index. All field values will be returned "in their zero-terminated string representation" (and accordingly an empty string for NULL values).

  - `myDatabase.Query(sql, RecordSet, ...)`
	Should be called for SELECT statements whenever you want to get the result of the query as a "RecordSet" object. You'll have to call the built-in method `RecordSet.Next()` to access the records sequentially. Only `DB-Query()` does handle BLOBs properly. All other field types will be returned as strings (see `DB.GetTable()`). If you don't need the RecordSet anymore, call `RecordSet.Free()` to release the resources.

  - `myDatabase.StoreBLOB(sql, BlobArray)`
	Should be called whenever BLOBs shall be stored in the database. For each BLOB in the row you have to specify a `?` parameter within the statement. The parameters are numbered automatically from left to right starting with 1. For each parameter you have to pass an object within BlobArray containing the address and the size of the BLOB.

- After all work is done, call `myDatabase.CloseDB()` to close the database. For all still existing queries `RecordSet.Free()` will be called internally.

- For further details look at the inline documentation in the class script and the sample scripts, please.
