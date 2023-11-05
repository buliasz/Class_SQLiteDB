; ======================================================================================================================
; Function:		Class definitions as wrappers for SQLite3.dll to work with SQLite DBs.
; AHK version:	2.0.10
; Tested on:	Win 11 (x64), SQLite 3.11.1
; Version:		0.0.01.00/2011-08-10/just me
;				0.0.02.00/2012-08-10/just me	-  Added basic BLOB support
;				0.0.03.00/2012-08-11/just me	-  Added more advanced BLOB support
;				0.0.04.00/2013-06-29/just me	-  Added new methods AttachDB and DetachDB
;				0.0.05.00/2013-08-03/just me	-  Changed base class assignment
;				0.0.06.00/2016-01-28/just me	-  Fixed version check, revised parameter initialization.
;				0.0.07.00/2016-03-28/just me	-  Added support for PRAGMA statements.
;				0.0.08.00/2019-03-09/just me	-  Added basic support for application-defined functions
;				0.0.09.00/2019-07-09/just me	-  Added basic support for prepared statements, minor bug fixes
;				0.0.10.00/2019-12-12/just me	-  Fixed bug in EscapeStr method
;				0.0.11.00/2021-10-10/just me	-  Removed statement checks in GetTable, Prepare, and Query
;				0.0.12.00/2022-09-18/just me	-  Fixed bug for Bind - type text
;				0.0.13.00/2022-10-03/just me	-  Fixed bug in Prepare
;				0.0.14.00/2022-10-04/just me	-  Changed DllCall parameter type PtrP to UPtrP
;				0.0.14.00/2022-11-04/buliasz	-  AHK 2 port and significant code refactor
; Remarks:		Names of "private" properties / methods are prefixed with an underscore,
;				they must not be set / called by the script!
;
;				SQLite3.dll file is assumed to be in the script's folder, otherwise you have to
;				provide an INI-File SQLiteDB.ini in the script's folder containing the path:
;				[Main]
;				DllPath=Path to SQLite3.dll
;
;				Encoding of SQLite DBs is assumed to be UTF-8
;				Minimum supported SQLite3.dll version is 3.6
;				Download the current version of SQLite3.dll (and also SQlite3.exe) from www.sqlite.org
; ======================================================================================================================
; This software is provided 'as-is', without any express or implied warranty.
; In no event will the authors be held liable for any damages arising from the
; use of this software.
; ======================================================================================================================

; ======================================================================================================================
; CLASS SQLiteClass - SQLite DB main class
; ======================================================================================================================
class SQLiteClass {
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; PUBLIC Interface ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	; ===================================================================================================================
	; Static Properties
	; ===================================================================================================================
	static version := ""
	static RET_CODE := {
		OK: 0,			  	; Successful result
		ERROR: 1,		  	; SQL error or missing database
		INTERNAL: 2,	  	; NOT USED. Internal logic error in SQLite
		PERM: 3,			; Access permission denied
		ABORT: 4,		  	; Callback routine requested an abort
		BUSY: 5,			; The database file is locked
		LOCKED: 6,			; A table in the database is locked
		NOMEM: 7,			; A malloc() failed
		READONLY: 8,		; Attempt to write a readonly database
		INTERRUPT: 9,		; Operation terminated by sqlite3_interrupt()
		IOERR: 10,			; Some kind of disk I/O error occurred
		CORRUPT: 11,		; The database disk image is malformed
		NOTFOUND: 12,		; NOT USED. Table or record not found
		FULL: 13,			; Insertion failed because database is full
		CANTOPEN: 14,		; Unable to open the database file
		PROTOCOL: 15,		; NOT USED. Database lock protocol error
		EMPTY: 16,			; Database is empty
		SCHEMA: 17,			; The database schema changed
		TOOBIG: 18,			; String or BLOB exceeds size limit
		CONSTRAINT: 19,		; Abort due to constraint violation
		MISMATCH: 20,		; Data type mismatch
		MISUSE: 21,			; Library used incorrectly
		NOLFS: 22,			; Uses OS features not supported on host
		AUTH: 23,			; Authorization denied
		FORMAT: 24,			; Auxiliary database format error
		RANGE: 25,			; 2nd parameter to sqlite3_bind out of range
		NOTADB: 26,			; File opened that is not a database file
		ROW: 100,			; sqlite3_step() has another row ready
		DONE: 101,			; sqlite3_step() has finished executing
	}
	static TYPES := {
		BLOB: 1,
		DOUBLE: 1,
		INT: 1,
		TEXT: 1,
	}

	; ===================================================================================================================
	; Properties
	; ===================================================================================================================
	errorMsg := ""				; Error message							(String)
	errorCode := 0				; SQLite error code						(Variant)
	changes := 0				; Changes made by last call of Exec()	(Integer)
	sql := ""					; Last executed SQL statement			(String)


	; ===================================================================================================================
	; METHOD OpenDB		Open a database
	; Parameters:		dbPath		- Path of the database file
	;					accessType  - Wanted access: "R"ead / "W"rite
	;					isCreate	- Create new database in write mode, if it doesn't exist
	; return values:	On success  - true
	;					On failure  - false, errorMsg / errorCode contain additional information
	; Remarks:			If dbPath is empty in write mode, a database called ":memory:" is created in memory
	;					and deleted on call of CloseDB.
	; ===================================================================================================================
	OpenDB(dbPath, accessType:="W", isCreate:=true) {
		static SQLITE_OPEN_READONLY  := 0x01	; Database opened as read-only
		static SQLITE_OPEN_READWRITE := 0x02	; Database opened as read-write
		static SQLITE_OPEN_CREATE	 := 0x04	; Database will be created if not exists
		static MEMDB := ":memory:"

		this.errorMsg := ""
		this.errorCode := 0
		dbHandle := 0

		if (dbPath == "") {
			dbPath := MEMDB
		}

		if (dbPath == this._path) && (this._handle) {
			return true
		}

		if (this._handle) {
			this.errorMsg := "You must first close DB " . this._path . "!"
			return false
		}

		flags := 0
		accessType := SubStr(accessType, 1, 1)
		if (accessType != "W") && (accessType != "R") {
			accessType := "R"
		}
		flags := SQLITE_OPEN_READONLY
		if (accessType == "W") {
			flags := SQLITE_OPEN_READWRITE
			if (isCreate) {
				flags |= SQLITE_OPEN_CREATE
			}
		}
		this._path := dbPath
		utf8 := this._StrToUtf8(dbPath)
		returnCode := DllCall("SQlite3.dll\sqlite3_open_v2", "Ptr", utf8, "UPtrP", &dbHandle, "Int", flags, "Ptr", 0, "Cdecl Int")
		if (returnCode) {
			this._path := ""
			this.errorMsg := this._ErrMsg()
			this.errorCode := returnCode
			return false
		}

		this._handle := dbHandle
		return true
	}

	; ===================================================================================================================
	; METHOD CloseDB		Close database
	; Parameters:			None
	; return values:		On success	- true
	;						On failure	- false, errorMsg / errorCode contain additional information
	; ===================================================================================================================
	CloseDB() {
		this.errorMsg := ""
		this.errorCode := 0
		this.sql := ""

		if (!this._handle) {
			return true
		}

		for query in this._queries {
			DllCall("SQlite3.dll\sqlite3_finalize", "Ptr", &query, "Cdecl Int")
		}

		returnCode := DllCall("SQlite3.dll\sqlite3_close", "Ptr", this._handle, "Cdecl Int")
		if (returnCode) {
			this.errorMsg := this._ErrMsg()
			this.errorCode := returnCode
			return false
		}

		this._path := ""
		this._handle := ""
		this._queries := []
		return true
	}

	; ===================================================================================================================
	; METHOD AttachDB	Add another database file to the current database connection
	;					http://www.sqlite.org/lang_attach.html
	; Parameters:		dbPath		- Path of the database file
	;					dbAlias		- Database alias name used internally by SQLite
	; return values:	On success  - true
	;					On failure	- false, errorMsg / errorCode contain additional information
	; ===================================================================================================================
	AttachDB(dbPath, dbAlias) {
		return this.Exec("ATTACH DATABASE '" . dbPath . "' As " . dbAlias . ";")
	}

	; ===================================================================================================================
	; METHOD DetachDB	Detaches an additional database connection previously attached using AttachDB()
	;					http://www.sqlite.org/lang_detach.html
	; Parameters:		dbAlias	  	- Database alias name used with AttachDB()
	; return values:	On success  - true
	;					On failure  - false, errorMsg / errorCode contain additional information
	; ===================================================================================================================
	DetachDB(dbAlias) {
		return this.Exec("DETACH DATABASE " . dbAlias . ";")
	}

	; ===================================================================================================================
	; METHOD Exec		Execute SQL statement
	; Parameters:		SQL			- Valid SQL statement
	;					Callback	- Name of a callback function to invoke for each result row coming out
	;								of the evaluated SQL statements.
	;								The function must accept 4 parameters:
	;								1: SQLiteClass object
	;								2: Number of columns
	;								3: Pointer to an array of pointers to columns text
	;								4: Pointer to an array of pointers to column names
	;								The address of the current SQL string is passed in A_EventInfo.
	;								if the callback function returns non-zero, DB.Exec() returns SQLITE_ABORT
	;								without invoking the callback again and without running any subsequent
	;								SQL statements.
	; return values:	On success	- true, the number of changed rows is given in property changes
	;					On failure	- false, errorMsg / errorCode contain additional information
	; ===================================================================================================================
	Exec(sql, Callback:="") {
		this.errorMsg := ""
		this.errorCode := 0
		this.sql := sql
		if (!this._handle) {
			this.errorMsg := "Invalid database handle!"
			return false
		}

		cbPtr := 0
		err := 0

		if (Callback && Callback.MinParams == 4) {
			cbPtr := CallbackCreate(Callback, "F C", 4)
		}
		utf8 := this._StrToUtf8(sql)
		objAddr := ObjPtrAddRef(this)
		returnCode := DllCall("SQlite3.dll\sqlite3_exec", "Ptr", this._handle, "Ptr", utf8, "Ptr", cbPtr, "Ptr", objAddr
						, "UPtrP", &err, "Cdecl Int")
		ObjRelease(objAddr)
		if (cbPtr) {
			CallbackFree(cbPtr)
		}

		if (returnCode) {
			this.errorMsg := StrGet(err, "UTF-8")
			this.errorCode := returnCode
			DllCall("SQLite3.dll\sqlite3_free", "Ptr", &err, "Cdecl")
			return false
		}

		this.changes := this._Changes()
		return true
	}

	; ===================================================================================================================
	; METHOD GetTable		Get complete result for SELECT query
	; Parameters:			sql	- SQL SELECT statement
	;						&TB	- Variable to store the result object (TB _Table)
	;							  maxResult	- Number of rows to return:
	;								  0			Complete result (default)
	;								 -1			return only rowCount and columnCount
	;								 -2			return counters and array columnNames
	;								  n			return counters and columnNames and first n rows
	; return values:		On success  - true, TB contains the result object
	;						On failure  - false, errorMsg / errorCode contain additional information
	; ===================================================================================================================
	GetTable(sql, &TB, maxResult:=0) {
		TB := ""
		this.errorMsg := ""
		this.errorCode := 0
		this.sql := sql
		if (!this._handle) {
			this.errorMsg := "Invalid database handle!"
			return false
		}

		names := ""
		err := 0, returnCode := 0, GetRows := 0
		I := 0, Rows := Cols := 0
		Table := 0
		if (!maxResult is Integer) {
			maxResult := 0
		}
		if (maxResult < -2) {
			maxResult := 0
		}
		utf8 := this._StrToUtf8(sql)
		returnCode := DllCall("SQlite3.dll\sqlite3_get_table", "Ptr", this._handle, "Ptr", utf8, "UPtrP", &Table
						, "IntP", Rows, "IntP", Cols, "UPtrP", &err, "Cdecl Int")
		if (returnCode) {
			this.errorMsg := StrGet(err, "UTF-8")
			this.errorCode := returnCode
			DllCall("SQLite3.dll\sqlite3_free", "Ptr", &err, "Cdecl")
			return false
		}

		TB := this._Table()
		TB.ColumnCount := Cols
		TB.rowCount := Rows
		if (maxResult == -1) {
			DllCall("SQLite3.dll\sqlite3_free_table", "Ptr", &Table, "Cdecl")
			return true
		}

		if (maxResult == -2) {
			GetRows := 0
		} else if (maxResult > 0 && maxResult <= Rows) {
			GetRows := maxResult
		} else {
			GetRows := Rows
		}

		Offset := 0
		names := Array()
		loop Cols {
			names[A_Index] := StrGet(NumGet(Table+0, Offset, "UPtr"), "UTF-8")
			Offset += A_PtrSize
		}

		TB.columnNames := names
		TB.hasNames := true
		loop GetRows {
			i := A_Index
			TB.rows[i] := []
			loop Cols {
				TB.rows[i][A_Index] := StrGet(NumGet(Table+0, Offset, "UPtr"), "UTF-8")
				Offset += A_PtrSize
			}
		}

		if (GetRows) {
			TB.hasRows := true
		}
		DllCall("SQLite3.dll\sqlite3_free_table", "Ptr", &Table, "Cdecl")

		return true
	}

	; ===================================================================================================================
	; Prepared statement 10:54 2019.07.05. by Dixtroy
	;  DB := SQLiteDB()
	;  DB.OpenDB(DBFileName)
	;  DB.Prepare 1 or more, just once
	;  DB.Step 1 or more on prepared one, repeatable
	;  DB.Finalize at the end
	; ===================================================================================================================
	; ===================================================================================================================
	; METHOD Prepare		Prepare database table for further actions.
	; Parameters:			sql			- SQL statement to be compiled
	;						&statement	- Variable to store the statement object (Class _Statement)
	; return values:		On success 	- true, ST contains the statement object
	;						On failure 	- false, errorMsg / errorCode contain additional information
	; Remarks:				You have to pass one ? for each column you want to assign a value later.
	; ===================================================================================================================
	Prepare(sql, &statement) {
		this.errorMsg := ""
		this.errorCode := 0
		this.sql := sql
		if (!this._handle) {
			this.errorMsg := "Invalid database handle!"
			return false
		}

		stmt := 0
		utf8 := this._StrToUtf8(sql)
		returnCode := DllCall("SQlite3.dll\sqlite3_prepare_v2", "Ptr", this._handle, "Ptr", utf8, "Int", -1
						, "UPtrP", &stmt, "Ptr", 0, "Cdecl Int")
		if (returnCode) {
			this.errorMsg := A_ThisFunc . ": " . this._ErrMsg()
			this.errorCode := returnCode
			return false
		}

		statement := this._Statement()
		statement.paramCount := DllCall("SQlite3.dll\sqlite3_bind_parameter_count", "Ptr", &stmt, "Cdecl Int")
		statement._handle := stmt
		statement._db := this
		this._stmts.Push(stmt)
		return true
	 }

	; ===================================================================================================================
	; METHOD Query		Get "recordset" object for prepared SELECT query
	; Parameters:		sql			- SQL SELECT statement
	;					&RS	 		- Variable to store the result object (Class _RecordSet)
	; return values:	On success  - true, RS contains the result object
	;					On failure  - false, errorMsg / errorCode contain additional information
	; ===================================================================================================================
	Query(sql, &RS) {
		RS := ""
		this.errorMsg := ""
		this.errorCode := 0
		this.sql := sql
		columnCount := 0
		hasRows := false
		if (!this._handle) {
			this.errorMsg := "Invalid dadabase handle!"
			return false
		}

		query := 0
		utf8 := this._StrToUtf8(sql)
		returnCode := DllCall("SQlite3.dll\sqlite3_prepare_v2", "Ptr", this._handle, "Ptr", utf8, "Int", -1
						, "UPtrP", &query, "Ptr", 0, "Cdecl Int")
		if (returnCode) {
			this.errorMsg := this._ErrMsg()
			this.errorCode := returnCode
			return false
		}

		returnCode := DllCall("SQlite3.dll\sqlite3_column_count", "Ptr", &query, "Cdecl Int")
		if (returnCode < 1) {
			this.errorMsg := "Query result is empty!"
			this.errorCode := SQLiteClass.RET_CODE.EMPTY
			return false
		}

		columnCount := returnCode
		names := []
		loop returnCode {
			namePtr := DllCall("SQlite3.dll\sqlite3_column_name", "Ptr", &query, "Int", A_Index - 1, "Cdecl UPtr")
			names[A_Index] := StrGet(namePtr, "UTF-8")
		}

		returnCode := DllCall("SQlite3.dll\sqlite3_step", "Ptr", &query, "Cdecl Int")
		if (returnCode == SQLiteClass.RET_CODE.ROW) {
			hasRows := true
		}
		returnCode := DllCall("SQlite3.dll\sqlite3_reset", "Ptr", &query, "Cdecl Int")
		RS := this._RecordSet()
		RS.ColumnCount := columnCount
		RS.columnNames := names
		RS.hasNames := true
		RS.hasRows := hasRows
		RS._handle := query
		RS._db := this
		this._queries.Push(query)
		return true
	}

	; ===================================================================================================================
	; METHOD CreateScalarFunc  Create a scalar application defined function
	; Parameters:		name 		-  the name of the function
	;					args  		-  the number of arguments that the SQL function takes
	;					Func  		-  a pointer to AHK functions that implement the SQL function
	;					enc			-  specifies what text encoding this SQL function prefers for its parameters
	;					param 		-  an arbitrary pointer accessible within the funtion with sqlite3_user_data()
	; return values:	On success 	- true
	;					On failure  - false, errorMsg / errorCode contain additional information
	; Documentation:	www.sqlite.org/c3ref/create_function.html
	; ===================================================================================================================
	CreateScalarFunc(name, args, Func, enc:=0x0801, param:=0) {
		; SQLITE_DETERMINISTIC == 0x0800 - the function will always return the same result given the same inputs
		;											within a single SQL statement
		; SQLITE_UTF8 == 0x0001
		this.errorMsg := ""
		this.errorCode := 0
		if (!this._handle) {
			this.errorMsg := "Invalid database handle!"
			return false
		}

		returnCode := DllCall("SQLite3.dll\sqlite3_create_function", "Ptr", this._handle, "AStr", name, "Int", args,
								"Int", &enc, "Ptr", &param, "Ptr", &Func, "Ptr", 0, "Ptr", 0, "Cdecl Int")
		if (returnCode) {
			this.errorMsg := this._ErrMsg()
			this.errorCode := returnCode
			return false
		}

		return true
	}

	; ===================================================================================================================
	; METHOD LastInsertRowID	Get the ROWID of the last inserted row
	; Parameters:		&rowID 		- Variable to store the ROWID
	; return values:	On success  - true, RowID contains the ROWID
	;					On failure  - false, errorMsg / errorCode contain additional information
	; ===================================================================================================================
	LastInsertRowID(&rowID) {
		this.errorMsg := ""
		this.errorCode := 0
		this.sql := ""
		if (!this._handle) {
			this.errorMsg := "Invalid database handle!"
			return false
		}
		rowID := 0
		returnCode := DllCall("SQLite3.dll\sqlite3_last_insert_rowid", "Ptr", this._handle, "Cdecl Int64")
		rowID := returnCode
		return true
	}

	; ===================================================================================================================
	; METHOD TotalChanges	Get the number of changed rows since connecting to the database
	; Parameters:		&rows  		- Variable to store the number of rows
	; return values:	On success  - true, Rows contains the number of rows
	;					On failure  - false, errorMsg / errorCode contain additional information
	; ===================================================================================================================
	TotalChanges(&rows) {
		this.errorMsg := ""
		this.errorCode := 0
		this.sql := ""
		if (!this._handle) {
			this.errorMsg := "Invalid database handle!"
			return false
		}
		rows := 0
		returnCode := DllCall("SQLite3.dll\sqlite3_total_changes", "Ptr", this._handle, "Cdecl Int")
		rows := returnCode
		return true
	}

	; ===================================================================================================================
	; METHOD SetTimeout	Set the timeout to wait before SQLITE_BUSY or SQLITE_IOERR_BLOCKED is returned,
	;					when a table is locked.
	; Parameters:		timeout  	- Time to wait in milliseconds
	; return values:	On success  - true
	;					On failure  - false, errorMsg / errorCode contain additional information
	; ===================================================================================================================
	SetTimeout(timeout:=1000) {
		this.errorMsg := ""
		this.errorCode := 0
		this.sql := ""
		if (!this._handle) {
			this.errorMsg := "Invalid database handle!"
			return false
		}
		if (!timeout is Integer) {
			timeout := 1000
		}
		returnCode := DllCall("SQLite3.dll\sqlite3_busy_timeout", "Ptr", this._handle, "Int", timeout, "Cdecl Int")
		if (returnCode) {
			this.errorMsg := this._ErrMsg()
			this.errorCode := returnCode
			return false
		}
		return true
	}

	; ===================================================================================================================
	; METHOD EscapeStr	Escapes special characters in a string to be used as field content
	; Parameters:		str			- String to be escaped
	;					quote		- Add single quotes around the outside of the total string (true / false)
	; return values:	On success  - true
	;					On failure  - false, errorMsg / errorCode contain additional information
	; ===================================================================================================================
	EscapeStr(&str, quote:=true) {
		this.errorMsg := ""
		this.errorCode := 0
		this.sql := ""

		if (!this._handle) {
			this.errorMsg := "Invalid database handle!"
			return false
		}

		if (Str is Number) {
			return true
		}

		op := Buffer(16, 0)
		StrPut(quote ? "%Q" : "%q", &op, "UTF-8")
		utf8 := this._StrToUtf8(str)
		ptr := DllCall("SQLite3.dll\sqlite3_mprintf", "Ptr", &op, "Ptr", utf8, "Cdecl UPtr")
		str := this._Utf8ToStr(ptr)
		DllCall("SQLite3.dll\sqlite3_free", "Ptr", ptr, "Cdecl")

		return true
	}

	; ===================================================================================================================
	; METHOD StoreBLOB	Use BLOBs as parameters of an INSERT/UPDATE/REPLACE statement.
	; Parameters:		sql			- SQL statement to be compiled
	;					blobArray	- Array of objects containing two keys/value pairs:
	; return values:	On success  - true
	;					On failure  - false, errorMsg / errorCode contain additional information
	; Remarks:			For each BLOB in the row you have to specify a ? parameter within the statement. The
	;					parameters are numbered automatically from left to right starting with 1.
	;					For each parameter you have to pass an object within 'blobArray' containing the address
	;					and the size of the BLOB.
	; ===================================================================================================================
	StoreBLOB(sql, blobArray) {
		static SQLITE_STATIC := 0
		this.errorMsg := ""
		this.errorCode := 0

		if (!this._handle) {
			this.errorMsg := "Invalid database handle!"
			return false
		}

		if (!RegExMatch(sql, "i)^\s*(INSERT|UPDATE|REPLACE)\s")) {
			this.errorMsg := A_ThisFunc . " requires an INSERT/UPDATE/REPLACE statement!"
			return false
		}

		query := 0
		utf8 := this._StrToUtf8(sql)
		returnCode := DllCall("SQlite3.dll\sqlite3_prepare_v2", "Ptr", this._handle, "Ptr", utf8, "Int", -1
						, "UPtrP", &query, "Ptr", 0, "Cdecl Int")
		if (returnCode) {
			this.errorMsg := A_ThisFunc . ": " . this._ErrMsg()
			this.errorCode := returnCode
			return false
		}

		for blobNum, blob In blobArray {
			if (!blob.address || !blob.Size) {
				this.errorMsg := A_ThisFunc . ": Invalid parameter blobArray!"
				this.errorCode := -1
				return false
			}

			returnCode := DllCall("SQlite3.dll\sqlite3_bind_blob", "Ptr", &query, "Int", blobNum, "Ptr", blob.address
							, "Int", blob.Size, "Ptr", SQLITE_STATIC, "Cdecl Int")
			if (returnCode) {
				this.errorMsg := A_ThisFunc . ": " . this._ErrMsg()
				this.errorCode := returnCode
				return false
			}
		}

		returnCode := DllCall("SQlite3.dll\sqlite3_step", "Ptr", &query, "Cdecl Int")
		if (returnCode) && (returnCode != SQLiteClass.RET_CODE.DONE) {
			this.errorMsg := A_ThisFunc . ": " . this._ErrMsg()
			this.errorCode := returnCode
			return false
		}

		returnCode := DllCall("SQlite3.dll\sqlite3_finalize", "Ptr", &query, "Cdecl Int")
		if (returnCode) {
			this.errorMsg := A_ThisFunc . ": " . this._ErrMsg()
			this.errorCode := returnCode
			return false
		}

		return true
	}

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; PRIVATE Properties and Methods ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	static _sqliteDll := A_ScriptDir . "\SQLite3.dll"
	static _refCount := 0
	static _minVersion := "3.6"

	; ===================================================================================================================
	; CLASS _Table
	; Object returned from method GetTable()
	; _Table is an independent object and does not need SQLite after creation at all.
	; ===================================================================================================================
	class _Table {
		; ----------------------------------------------------------------------------------------------------------------
		; CONSTRUCTOR  Create instance variables
		; ----------------------------------------------------------------------------------------------------------------
		__New() {
			 this.columnCount := 0			; Number of columns in the result table		(Integer)
			 this.rowCount := 0				; Number of rows in the result table			(Integer)
			 this.columnNames := []			; Names of columns in the result table		(Array)
			 this.rows := []				; Rows of the result table					(Array of Arrays)
			 this.hasNames := false			; Does var columnNames contain names?		(Bool)
			 this.hasRows := false			; Does var Rows contain rows?				(Bool)
			 this._currentRow := 0			; row index of last returned row				(Integer)
		}

		; ----------------------------------------------------------------------------------------------------------------
		; METHOD GetRow		Get row for RowIndex
		; Parameters:		rowIndex	- Index of the row to retrieve, the index of the first row is 1
		;					&row		- Variable to pass out the row array
		; return values:	On failure  - false
		;					On success  - true, row contains a valid array
		; Remarks:			_currentRow is set to RowIndex, so a subsequent call of NextRow() will return the
		;					following row.
		; ----------------------------------------------------------------------------------------------------------------
		GetRow(rowIndex, &row) {
			row := ""
			if (rowIndex < 1 || rowIndex > this.rowCount) {
				return false
			}
			if (!this.rows.HasKey(rowIndex)) {
				return false
			}

			row := this.rows[rowIndex]
			this._currentRow := rowIndex

			return true
		}

		; ----------------------------------------------------------------------------------------------------------------
		; METHOD Next		Get next row depending on _currentRow
		; Parameters:		&Row		- Variable to pass out the row array
		; return values:	On failure  - false, -1 for EOR (end of rows)
		;					On success  - true, row contains a valid array
		; ----------------------------------------------------------------------------------------------------------------
		Next(&row) {
			row := ""

			if (this._currentRow >= this.rowCount) {
				return -1
			}

			this._currentRow += 1
			if (!this.rows.HasKey(this._currentRow)) {
				return false
			}

			row := this.rows[this._currentRow]
			return true
		}

		; ----------------------------------------------------------------------------------------------------------------
		; METHOD Reset	Reset _currentRow to zero
		; Parameters:	None
		; return value:	true
		; ----------------------------------------------------------------------------------------------------------------
		Reset() {
			this._currentRow := 0
			return true
		}
	}

	; ===================================================================================================================
	; CLASS _RecordSet
	; Object returned from method Query()
	; The records (rows) of a recordset can be accessed sequentially per call of Next() starting with the first record.
	; After a call of Reset() calls of Next() will start with the first record again.
	; When the recordset isn't needed any more, call Free() to free the resources.
	; The lifetime of a recordset depends on the lifetime of the related SQLiteClass object.
	; ===================================================================================================================
	class _RecordSet {
		; ----------------------------------------------------------------------------------------------------------------
		; CONSTRUCTOR  Create instance variables
		; ----------------------------------------------------------------------------------------------------------------
		__New() {
			this.columnCount := 0		; Number of columns						(Integer)
			this.columnNames := []		; Names of columns in the result table	(Array)
			this.hasNames := false		; Does var columnNames contain names?	(Bool)
			this.hasRows := false		; Does _RecordSet contain rows?			(Bool)
			this.currentRow := 0		; Index of current row					(Integer)
			this.errorMsg := ""			; Last error message					(String)
			this.errorCode := 0			; Last SQLite error code				(Variant)
			this._handle := 0			; Query handle							(Pointer)
			this._db := {}				; SQLiteClass object					(Object)
		}

		; ----------------------------------------------------------------------------------------------------------------
		; DESTRUCTOR	Clear instance variables
		; ----------------------------------------------------------------------------------------------------------------
		__Delete() {
			if (this._handle)
				this.Free()
		}

		; ----------------------------------------------------------------------------------------------------------------
		; METHOD Next		Get next row of query result
		; Parameters:		&Row		- Variable to store the row array
		; return values:	On success  - true, row contains the row array
		;					On failure  - false, errorMsg / errorCode contain additional information
		;					-1 for EOR (end of records)
		; ----------------------------------------------------------------------------------------------------------------
		Next(&row) {
			static SQLITE_NULL := 5
			static SQLITE_BLOB := 4
			static EOR := -1

			row := ""
			this.errorMsg := ""
			this.errorCode := 0
			if (!this._handle) {
				this.errorMsg := "Invalid query handle!"
				return false
			}

			returnCode := DllCall("SQlite3.dll\sqlite3_step", "Ptr", this._handle, "Cdecl Int")
			if (returnCode != SQLiteClass.RET_CODE.ROW) {
				if (returnCode == SQLiteClass.RET_CODE.DONE) {
					this.errorMsg := "EOR"
					this.errorCode := returnCode
					return EOR
				}

				this.errorMsg := this._db.ErrMsg()
				this.errorCode := returnCode
				return false
			}

			returnCode := DllCall("SQlite3.dll\sqlite3_data_count", "Ptr", this._handle, "Cdecl Int")
			if (returnCode < 1) {
				this.errorMsg := "Recordset is empty!"
				this.errorCode := SQLiteClass.RET_CODE.EMPTY
				return false
			}

			row := []
			loop returnCode {
				column := A_Index - 1
				columnType := DllCall("SQlite3.dll\sqlite3_column_type", "Ptr", this._handle, "Int", column, "Cdecl Int")

				if (columnType == SQLITE_NULL) {
					row[A_Index] := ""
				} else if (columnType == SQLITE_BLOB) {
					blobPtr := DllCall("SQlite3.dll\sqlite3_column_blob", "Ptr", this._handle, "Int", column, "Cdecl UPtr")
					blobSize := DllCall("SQlite3.dll\sqlite3_column_bytes", "Ptr", this._handle, "Int", column, "Cdecl Int")
					if (blobPtr == 0) || (blobSize == 0) {
						row[A_Index] := ""
					} else {
						row[A_Index] := {}
						row[A_Index].Size := blobSize
						row[A_Index].blob := Buffer(blobSize)
						address := row[A_Index].blob.Ptr
						DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", address, "Ptr", blobPtr, "Ptr", blobSize)
					}
				} else {
					textPtr := DllCall("SQlite3.dll\sqlite3_column_text", "Ptr", this._handle, "Int", column, "Cdecl UPtr")
					row[A_Index] := StrGet(textPtr, "UTF-8")
				}
			}

			this.currentRow += 1
			return true
		}

		; ----------------------------------------------------------------------------------------------------------------
		; METHOD Reset		Reset the result pointer
		; Parameters:		None
		; return values:	On success  - true
		;					On failure  - false, errorMsg / errorCode contain additional information
		; Remarks:			After a call of this method you can access the query result via Next() again.
		; ----------------------------------------------------------------------------------------------------------------
		Reset() {
			this.errorMsg := ""
			this.errorCode := 0
			if (!this._handle) {
				this.errorMsg := "Invalid query handle!"
				return false
			}

			returnCode := DllCall("SQlite3.dll\sqlite3_reset", "Ptr", this._handle, "Cdecl Int")
			if (returnCode) {
				this.errorMsg := this._db._ErrMsg()
				this.errorCode := returnCode
				return false
			}

			this.currentRow := 0
			return true
		}

		; ----------------------------------------------------------------------------------------------------------------
		; METHOD Free		Free query result
		; Parameters:		None
		; return values:	On success  - true
		;					On failure  - false, errorMsg / errorCode contain additional information
		; Remarks:			After the call of this method further access on the query result is impossible.
		; ----------------------------------------------------------------------------------------------------------------
		Free() {
			this.errorMsg := ""
			this.errorCode := 0
			if (!this._handle) {
				return true
			}

			returnCode := DllCall("SQlite3.dll\sqlite3_finalize", "Ptr", this._handle, "Cdecl Int")
			if (returnCode) {
				this.errorMsg := this._db._ErrMsg()
				this.errorCode := returnCode
				return false
			}

			this._handle := 0
			this._db := 0
			return true
		}
	}

	; ===================================================================================================================
	; CLASS _Statement
	; Object returned from method Prepare()
	; The life-cycle of a prepared statement object usually goes like this:
	; 1. Create the prepared statement object (PST) by calling DB.Prepare().
	; 2. Bind values to parameters using the PST.Bind_*() methods of the statement object.
	; 3. Run the SQL by calling PST.Step() one or more times.
	; 4. Reset the prepared statement using PTS.Reset() then go back to step 2. Do this zero or more times.
	; 5. Destroy the object using PST.Finalize().
	; The lifetime of a prepared statement depends on the lifetime of the related SQLiteClass object.
	; ===================================================================================================================
	class _Statement {
		; ----------------------------------------------------------------------------------------------------------------
		; CONSTRUCTOR  Create instance variables
		; ----------------------------------------------------------------------------------------------------------------
		__New() {
			this.errorMsg := ""			; Last error message									 (String)
			this.errorCode := 0			; Last SQLite error code								(Variant)
			this.paramCount := 0			; Number of SQL parameters for this statement	(Integer)
			this._handle := 0				; Query handle											 (Pointer)
			this._db := {}					; SQLiteClass object										 (Object)
		}

		; ----------------------------------------------------------------------------------------------------------------
		; DESTRUCTOR	Clear instance variables
		; ----------------------------------------------------------------------------------------------------------------
		__Delete() {
			if (this._handle)
				this.Free()
		}

		; ----------------------------------------------------------------------------------------------------------------
		; METHOD Bind		Bind values to SQL parameters.
		; Parameters:		Index		- 1-based index of the SQL parameter
		;					Type	 	- type, one of the SQLiteClass.TYPES
		;					param3		- type dependent value
		;					param4		- type dependent value
		;					param5		- not used
		; return values:	On success  - true
		;					On failure  - false, errorMsg / errorCode contain additional information
		; ----------------------------------------------------------------------------------------------------------------
		Bind(index, type, param3:="", param4:=0, param5:=0) {
			this.errorMsg := ""
			this.errorCode := 0

			if (!this._handle) {
				this.errorMsg := "Invalid statement handle!"
				return false
			}

			if (index < 1) || (index > this.paramCount) {
				this.errorMsg := "Invalid parameter index!"
				return false
			}

			if (type == SQLiteClass.TYPES.BLOB) {	; ----------------------------------------------------------------------------------------
				; param3 == BLOB pointer, param4 == BLOB size in bytes
				if (!param3 is Integer) {
					this.errorMsg := "Invalid blob pointer!"
					return false
				}

				if (!param4 is Integer) {
					this.errorMsg := "Invalid blob size!"
					return false
				}

				; Let SQLite always create a copy of the BLOB
				returnCode := DllCall("SQlite3.dll\sqlite3_bind_blob", "Ptr", this._handle, "Int", index, "Ptr", &param3
								, "Int", param4, "Ptr", -1, "Cdecl Int")
				if (returnCode) {
					this.errorMsg := this._ErrMsg()
					this.errorCode := returnCode
					return false
				}
			}
			else if (type == SQLiteClass.TYPES.DOUBLE) {	; ---------------------------------------------------------------------------------
				; param3 == double value
				if (!param3 is Float) {
					this.errorMsg := "Invalid value for double!"
					return false
				}

				returnCode := DllCall("SQlite3.dll\sqlite3_bind_double", "Ptr", this._handle, "Int", index, "Double", param3
								, "Cdecl Int")
				if (returnCode) {
					this.errorMsg := this._ErrMsg()
					this.errorCode := returnCode
					return false
				}
			}
			else if (type == SQLiteClass.TYPES.INT) {	; ------------------------------------------------------------------------------------
				; param3 == integer value
				if (!param3 is Integer) {
					this.errorMsg := "Invalid value for int!"
					return false
				}

				returnCode := DllCall("SQlite3.dll\sqlite3_bind_int", "Ptr", this._handle, "Int", index, "Int", param3
								, "Cdecl Int")
				if (returnCode) {
					this.errorMsg := this._ErrMsg()
					this.errorCode := returnCode
					return false
				}
			}
			else if (type == SQLiteClass.TYPES.TEXT) {	; -----------------------------------------------------------------------------------
				; param3 == zero-terminated string
				utf8 := this._db._StrToUtf8(param3)
				; Let SQLite always create a copy of the text
				returnCode := DllCall("SQlite3.dll\sqlite3_bind_text", "Ptr", this._handle, "Int", index, "Ptr", utf8
								, "Int", -1, "Ptr", -1, "Cdecl Int")
				if (returnCode) {
					this.errorMsg := this._ErrMsg()
					this.errorCode := returnCode
					return false
				}
			}
			else {
				this.errorMsg := "Invalid parameter type=" type "!"
				return false
			}

			return true
		}

		; ----------------------------------------------------------------------------------------------------------------
		; METHOD Step		Evaluate the prepared statement.
		; Parameters:		None
		; return values:	On success  - true
		;					On failure  - false, errorMsg / errorCode contain additional information
		; Remarks:			You must call statement.Reset() before you can call statement.Step() again.
		; ----------------------------------------------------------------------------------------------------------------
		Step() {
			this.errorMsg := ""
			this.errorCode := 0
			if (!this._handle) {
				this.errorMsg := "Invalid statement handle!"
				return false
			}

			returnCode := DllCall("SQlite3.dll\sqlite3_step", "Ptr", this._handle, "Cdecl Int")
			if (returnCode != SQLiteClass.RET_CODE.DONE
					&& returnCode != SQLiteClass.RET_CODE.ROW
			) {
				this.errorMsg := this._db.ErrMsg()
				this.errorCode := returnCode
				return false
			}

			return true
		}

		; ----------------------------------------------------------------------------------------------------------------
		; METHOD Reset		Reset the prepared statement.
		; Parameters:		clearBindings  	- Clear bound SQL parameter values (true/False)
		; return values:	On success		- true
		;					On failure		- false, errorMsg / errorCode contain additional information
		; Remarks:			After a call of this method you can access the query result via Next() again.
		; ----------------------------------------------------------------------------------------------------------------
		Reset(clearBindings:=true) {
			this.errorMsg := ""
			this.errorCode := 0
			if (!this._handle) {
				this.errorMsg := "Invalid statement handle!"
				return false
			}

			returnCode := DllCall("SQlite3.dll\sqlite3_reset", "Ptr", this._handle, "Cdecl Int")
			if (returnCode) {
				this.errorMsg := this._db._ErrMsg()
				this.errorCode := returnCode
				return false
			}

			if (clearBindings) {
				returnCode := DllCall("SQlite3.dll\sqlite3_clear_bindings", "Ptr", this._handle, "Cdecl Int")
				if (returnCode) {
					this.errorMsg := this._db._ErrMsg()
					this.errorCode := returnCode
					return false
				}
			}

			return true
		}

		; ----------------------------------------------------------------------------------------------------------------
		; METHOD Free		Free the prepared statement object.
		; Parameters:		None
		; return values:	On success  - true
		;					On failure  - false, errorMsg / errorCode contain additional information
		; Remarks:			After the call of this method further access on the statement object is impossible.
		; ----------------------------------------------------------------------------------------------------------------
		Free() {
			this.errorMsg := ""
			this.errorCode := 0
			if (!this._handle) {
				return true
			}

			returnCode := DllCall("SQlite3.dll\sqlite3_finalize", "Ptr", this._handle, "Cdecl Int")
			if (returnCode) {
				this.errorMsg := this._db._ErrMsg()
				this.errorCode := returnCode
				return false
			}

			this._handle := 0
			this._db := 0
			return true
		}
	}

	; ===================================================================================================================
	; CONSTRUCTOR __New
	; ===================================================================================================================
	__New() {
		this._path := ""				; Database path				(String)
		this._handle := 0				; Database handle			(Pointer)
		this._queries := []				; Valid queries				(List)
		this._stmts := []				; Valid prepared statements	(List)

		if (SQLiteClass._refCount == 0) {
			sqliteDll := SQLiteClass._sqliteDll
			if (!FileExist(sqliteDll) && FileExist(A_ScriptDir . "\SQLiteDB.ini")) {
				sqliteDll := IniRead(sqliteDll, A_ScriptDir "\SQLiteDB.ini", "Main", "DllPath")
				SQLiteClass._sqliteDll := sqliteDll
			}

			dll := DllCall("LoadLibrary", "Str", SQLiteClass._sqliteDll, "UPtr")
			if (!dll) {
				MsgBox("SQLiteDB Error, DLL " sqliteDll " does not exist!")
				ExitApp
			}

			SQLiteClass.version := StrGet(DllCall("SQlite3.dll\sqlite3_libversion", "Cdecl UPtr"), "UTF-8")
			sqlVersion := StrSplit(SQLiteClass.version, ".")
			minVersion := StrSplit(SQLiteClass._minVersion, ".")
			if (sqlVersion[1] < minVersion[1]) || ((sqlVersion[1] == minVersion[1]) && (sqlVersion[2] < minVersion[2])){
				DllCall("FreeLibrary", "Ptr", &dll)
				MsgBox("SQLite ERROR, Version " SQLiteClass.version " of SQLite3.dll is not supported!`n`n"
												  . "You can download the current version from www.sqlite.org!")
				ExitApp
			}
		}

		SQLiteClass._refCount += 1
	}

	; ===================================================================================================================
	; DESTRUCTOR __Delete
	; ===================================================================================================================
	__Delete() {
		if (this._handle) {
			this.CloseDB()
		}

		SQLiteClass._refCount -= 1
		if (SQLiteClass._refCount == 0) {
			dll := DllCall("GetModuleHandle", "Str", SQLiteClass._sqliteDll, "UPtr")
			if (dll) {
				DllCall("FreeLibrary", "Ptr", dll)
			}
		}
	}

	; ===================================================================================================================
	; PRIVATE _StrToUtf8
	; ===================================================================================================================
	_StrToUtf8(str) {
		utf8 := Buffer(StrPut(str, "UTF-8"), 0)
		StrPut(str, utf8, "UTF-8")
		return utf8
	}

	; ===================================================================================================================
	; PRIVATE _Utf8ToStr
	; ===================================================================================================================
	_Utf8ToStr(utf8) {
		return StrGet(utf8, "UTF-8")
	}

	; ===================================================================================================================
	; PRIVATE _ErrMsg
	; ===================================================================================================================
	_ErrMsg() {
		returnCode := DllCall("SQLite3.dll\sqlite3_errmsg", "Ptr", this._handle, "Cdecl UPtr")
		if (returnCode) {
			return StrGet(&returnCode, "UTF-8")
		}
		return ""
	}

	; ===================================================================================================================
	; PRIVATE _ErrCode
	; ===================================================================================================================
	_ErrCode() {
		return DllCall("SQLite3.dll\sqlite3_errcode", "Ptr", this._handle, "Cdecl Int")
	}

	; ===================================================================================================================
	; PRIVATE _Changes
	; ===================================================================================================================
	_Changes() {
		return DllCall("SQLite3.dll\sqlite3_changes", "Ptr", this._handle, "Cdecl Int")
	}
}

; ======================================================================================================================
; Exemplary custom callback function regexp()
; Parameters:		context -  handle to a sqlite3_context object
;					argC	-  number of elements passed in values (must be 2 for this function)
;					values	-  pointer to an array of pointers which can be passed to sqlite3_value_text():
;								1. Needle
;								2. Haystack
; return values:	Call sqlite3_result_int() passing 1 (true) for a match, otherwise pass 0 (False).
; ======================================================================================================================
SQLiteDB_RegExp(context, argC, values) {
	result := 0
	if (argC == 2) {
		addressN := DllCall("SQLite3.dll\sqlite3_value_text", "Ptr", NumGet(values + 0, "UPtr"), "Cdecl UPtr")
		addressH := DllCall("SQLite3.dll\sqlite3_value_text", "Ptr", NumGet(values + A_PtrSize, "UPtr"), "Cdecl UPtr")
		result := RegExMatch(StrGet(addressH, "UTF-8"), StrGet(addressN, "UTF-8"))
	}
	DllCall("SQLite3.dll\sqlite3_result_int", "Ptr", context, "Int", !!result, "Cdecl")	; 0 == false, 1 == trus
}
