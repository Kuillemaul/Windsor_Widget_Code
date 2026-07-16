/*
    Recover a supplier value after an incorrect duplicate-item deletion.

    Replace @ItemNumber and @Supplier with the confirmed values. The script
    automatically uses the first supplier field supported by Windsor Widget:
    supplier_name, supplier_code, Column1, or Supplier.
*/

SET XACT_ABORT ON;

DECLARE @ItemNumber NVARCHAR(100) = N'LFISB10105';
DECLARE @Supplier   NVARCHAR(255) = N'REPLACE WITH THE CORRECT SUPPLIER';
DECLARE @SupplierColumn SYSNAME;
DECLARE @Sql NVARCHAR(MAX);

IF @Supplier = N'REPLACE WITH THE CORRECT SUPPLIER'
    THROW 50001, 'Enter the confirmed supplier name before running this script.', 1;

SELECT TOP (1)
    @SupplierColumn = c.name
FROM sys.columns c
WHERE c.object_id = OBJECT_ID(N'dbo.items')
  AND c.name IN (N'supplier_name', N'supplier_code', N'Column1', N'Supplier')
ORDER BY CASE c.name
    WHEN N'supplier_name' THEN 1
    WHEN N'supplier_code' THEN 2
    WHEN N'Column1' THEN 3
    WHEN N'Supplier' THEN 4
    ELSE 99
END;

IF @SupplierColumn IS NULL
    THROW 50002, 'No supported supplier column exists in dbo.items.', 1;

BEGIN TRY
    BEGIN TRANSACTION;

    SET @Sql = N'
        UPDATE dbo.items
        SET ' + QUOTENAME(@SupplierColumn) + N' = @Supplier
        WHERE UPPER(LTRIM(RTRIM(item_number))) = UPPER(LTRIM(RTRIM(@ItemNumber)));

        IF @@ROWCOUNT <> 1
            THROW 50003, ''Expected exactly one item-master row to update.'', 1;
    ';

    EXEC sys.sp_executesql
        @Sql,
        N'@Supplier NVARCHAR(255), @ItemNumber NVARCHAR(100)',
        @Supplier = @Supplier,
        @ItemNumber = @ItemNumber;

    IF OBJECT_ID(N'dbo.supplier_master', N'U') IS NOT NULL
       AND NOT EXISTS (
            SELECT 1
            FROM dbo.supplier_master
            WHERE UPPER(LTRIM(RTRIM(supplier_name))) = UPPER(LTRIM(RTRIM(@Supplier)))
       )
    BEGIN
        INSERT INTO dbo.supplier_master (supplier_name)
        VALUES (@Supplier);
    END;

    SET @Sql = N'
        SELECT
            item_number,
            item_name,
            description,
            ' + QUOTENAME(@SupplierColumn) + N' AS supplier
        FROM dbo.items
        WHERE UPPER(LTRIM(RTRIM(item_number))) = UPPER(LTRIM(RTRIM(@ItemNumber)));
    ';

    EXEC sys.sp_executesql
        @Sql,
        N'@ItemNumber NVARCHAR(100)',
        @ItemNumber = @ItemNumber;

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;
    THROW;
END CATCH;
