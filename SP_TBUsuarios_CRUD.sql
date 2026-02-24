-- Si no existe, crear stub para poder ALTER sin errores
IF OBJECT_ID('dbo.SP_TBUsuarios_CRUD', 'P') IS NULL
    EXEC ('CREATE PROCEDURE dbo.SP_TBUsuarios_CRUD AS BEGIN SET NOCOUNT ON; END');
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

ALTER PROCEDURE dbo.SP_TBUsuarios_CRUD
    @opcion      CHAR(1),                
    @Cod_Usu     VARCHAR(4) = NULL,     
    @ApeNomUser  VARCHAR(50) = NULL,    
    @Password    VARCHAR(32) = NULL      
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    ----------------------------------------------------------------------
    -- Validar opción
    ----------------------------------------------------------------------
    IF @opcion NOT IN ('I','U','D','S')
    BEGIN
        RAISERROR('Opción inválida. Use: I (insertar), U (actualizar), D (eliminar), S (seleccionar).', 16, 1);
        RETURN;
    END

    ----------------------------------------------------------------------
    -- Helper: Normalizar @Cod_Usu (para U/D/S)
    ----------------------------------------------------------------------
    DECLARE @Cod_Usu_norm CHAR(4) = NULL;

    IF @Cod_Usu IS NOT NULL AND LTRIM(RTRIM(@Cod_Usu)) <> ''
    BEGIN
        DECLARE @CodTrim VARCHAR(50) = LTRIM(RTRIM(@Cod_Usu));
        -- Solo dígitos 0-9
        IF @CodTrim NOT LIKE '%[^0-9]%'
        BEGIN
            IF LEN(@CodTrim) > 4
            BEGIN
                RAISERROR('Cod_Usu numérico no puede exceder 4 caracteres.', 16, 1);
                RETURN;
            END
            SET @Cod_Usu_norm = RIGHT(REPLICATE('0', 4) + @CodTrim, 4);
        END
        ELSE
        BEGIN
            IF LEN(@CodTrim) <> 4
            BEGIN
                RAISERROR('Cod_Usu debe tener exactamente 4 caracteres cuando no es numérico.', 16, 1);
                RETURN;
            END
            SET @Cod_Usu_norm = @CodTrim;
        END
    END

    ----------------------------------------------------------------------
    -- INSERTAR (I)
    -- - IDUser: siguiente numérico → char(8) con ceros a la izquierda
    -- - Cod_Usu: primer hueco libre entre 0541 y 9999
    -- - Fechas: char(8) estilo 112 (yyyymmdd)
    ----------------------------------------------------------------------
    IF @opcion = 'I'
    BEGIN
        IF @ApeNomUser IS NULL OR LTRIM(RTRIM(@ApeNomUser)) = ''
        BEGIN
            RAISERROR('ApeNomUser es obligatorio y no puede ser vacío.', 16, 1);
            RETURN;
        END

        IF @Password IS NULL OR LTRIM(RTRIM(@Password)) = ''
        BEGIN
            RAISERROR('Password es obligatorio y no puede ser vacío.', 16, 1);
            RETURN;
        END

        DECLARE @NewIDInt     INT,
                @NewIDChar    CHAR(8),
                @NextCodInt   INT,
                @NewCodChar   CHAR(4),
                @now          DATETIME,
                @CodUsuFloor  INT;

        SET @now = GETDATE();
        SET @CodUsuFloor = 541;  -- <<< PISO: continuar desde 0541

        -- Validar rango del piso
        IF @CodUsuFloor IS NULL OR @CodUsuFloor < 1 OR @CodUsuFloor > 9999
        BEGIN
            RAISERROR('Valor inválido de piso para Cod_Usu. Debe estar entre 0001 y 9999.', 16, 1);
            RETURN;
        END

        -- Para evitar condiciones de carrera al buscar "primer hueco"
        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;
        BEGIN TRAN;

        -- Siguiente IDUser (tomando solo los IDUser numéricos)
        SELECT @NewIDInt = COALESCE(MAX(CONVERT(INT, IDUser)), 0) + 1
        FROM dbo.TBUsuarios
        WHERE IDUser NOT LIKE '%[^0-9]%';

        ------------------------------------------------------------------
        -- Buscar el primer código libre en [@CodUsuFloor .. 9999]
        -- Generamos la serie con ROW_NUMBER()
        ------------------------------------------------------------------
        ;WITH N AS
        (
            SELECT TOP (9999 - @CodUsuFloor + 1)
                   (ROW_NUMBER() OVER (ORDER BY (SELECT 1)) - 1) + @CodUsuFloor AS n
            FROM sys.all_objects
        )
        SELECT @NextCodInt = MIN(N.n)
        FROM N
        LEFT JOIN dbo.TBUsuarios U
          ON U.Cod_Usu NOT LIKE '%[^0-9]%'
         AND CONVERT(INT, U.Cod_Usu) = N.n
        WHERE U.Cod_Usu IS NULL;  -- hueco disponible

        IF @NextCodInt IS NULL
        BEGIN
            ROLLBACK TRAN;
            SET TRANSACTION ISOLATION LEVEL READ COMMITTED;
            DECLARE @rango VARCHAR(20) = RIGHT('0000' + CONVERT(VARCHAR(10), @CodUsuFloor), 4) + '-9999';
            RAISERROR('No hay códigos disponibles en el rango %s.', 16, 1, @rango);
            RETURN;
        END

        -- Validar que IDUser no desborde 8 dígitos
        IF LEN(CONVERT(VARCHAR(50), @NewIDInt)) > 8
        BEGIN
            ROLLBACK TRAN;
            SET TRANSACTION ISOLATION LEVEL READ COMMITTED;
            RAISERROR('El nuevo IDUser (%d) excede 8 caracteres.', 16, 1, @NewIDInt);
            RETURN;
        END

        -- Formateos (relleno con ceros)
        SET @NewIDChar  = RIGHT(REPLICATE('0', 8) + CONVERT(VARCHAR(8),  @NewIDInt), 8);
        SET @NewCodChar = RIGHT(REPLICATE('0', 4) + CONVERT(VARCHAR(4),  @NextCodInt), 4);

        BEGIN TRY
            INSERT INTO dbo.TBUsuarios
            (
                IDUser,       -- char(8)
                Cod_Usu,      -- char(4)
                ApeNomUser,   -- varchar(50)
                FecCreacion,  -- char(8) yyyymmdd
                FecIni,       -- char(8) yyyymmdd
                FecFin,       -- char(8) yyyymmdd
                [Password]    -- varchar(32)
            )
            VALUES
            (
                @NewIDChar,
                @NewCodChar,
                LTRIM(RTRIM(@ApeNomUser)),
                CONVERT(CHAR(8), @now, 112),
                CONVERT(CHAR(8), @now, 112),
                CONVERT(CHAR(8), @now, 112),
                @Password
            );

            COMMIT TRAN;
            SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

            SELECT TOP(1) *
            FROM dbo.TBUsuarios
            WHERE IDUser = @NewIDChar;
        END TRY
        BEGIN CATCH
            ROLLBACK TRAN;
            SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

            DECLARE @ErrMsg NVARCHAR(4000) = ERROR_MESSAGE();
            DECLARE @ErrNum INT = ERROR_NUMBER();
            DECLARE @ErrLine INT = ERROR_LINE();
            DECLARE @ErrProc SYSNAME = ERROR_PROCEDURE();
            DECLARE @ErrProcName SYSNAME = ISNULL(@ErrProc, 'SP_TBUsuarios_CRUD');

            RAISERROR('Error %d en %s (línea %d): %s', 16, 1, @ErrNum, @ErrProcName, @ErrLine, @ErrMsg);
        END CATCH

        RETURN;
    END

    ----------------------------------------------------------------------
    -- ACTUALIZAR (U)
    ----------------------------------------------------------------------
    IF @opcion = 'U'
    BEGIN
        IF @Cod_Usu_norm IS NULL
        BEGIN
            RAISERROR('Cod_Usu es obligatorio para actualizar.', 16, 1);
            RETURN;
        END

        IF @ApeNomUser IS NULL OR LTRIM(RTRIM(@ApeNomUser)) = ''
        BEGIN
            RAISERROR('ApeNomUser no puede ser NULL ni vacío en actualización.', 16, 1);
            RETURN;
        END

        IF @Password IS NULL OR LTRIM(RTRIM(@Password)) = ''
        BEGIN
            RAISERROR('Password no puede ser NULL ni vacío en actualización.', 16, 1);
            RETURN;
        END

        IF NOT EXISTS (SELECT 1 FROM dbo.TBUsuarios WHERE Cod_Usu = @Cod_Usu_norm)
        BEGIN
            RAISERROR('No existe un usuario con el Cod_Usu proporcionado.', 16, 1);
            RETURN;
        END

        BEGIN TRY
            UPDATE U
            SET U.ApeNomUser = LTRIM(RTRIM(@ApeNomUser)),
                U.[Password] = @Password
            FROM dbo.TBUsuarios AS U
            WHERE U.Cod_Usu = @Cod_Usu_norm;

            SELECT * FROM dbo.TBUsuarios WHERE Cod_Usu = @Cod_Usu_norm;
        END TRY
        BEGIN CATCH
            DECLARE @ErrMsg2 NVARCHAR(4000) = ERROR_MESSAGE();
            DECLARE @ErrNum2 INT = ERROR_NUMBER();
            DECLARE @ErrLine2 INT = ERROR_LINE();
            DECLARE @ErrProc2 SYSNAME = ERROR_PROCEDURE();
            DECLARE @ErrProcName2 SYSNAME = ISNULL(@ErrProc2, 'SP_TBUsuarios_CRUD');

            RAISERROR('Error %d en %s (línea %d): %s', 16, 1, @ErrNum2, @ErrProcName2, @ErrLine2, @ErrMsg2);
        END CATCH

        RETURN;
    END

    ----------------------------------------------------------------------
    -- ELIMINAR (D)
    ----------------------------------------------------------------------
    IF @opcion = 'D'
    BEGIN
        IF @Cod_Usu_norm IS NULL
        BEGIN
            RAISERROR('Cod_Usu es obligatorio para eliminar.', 16, 1);
            RETURN;
        END

        IF NOT EXISTS (SELECT 1 FROM dbo.TBUsuarios WHERE Cod_Usu = @Cod_Usu_norm)
        BEGIN
            RAISERROR('No existe un usuario con el Cod_Usu proporcionado para eliminar.', 16, 1);
            RETURN;
        END

        BEGIN TRY
            DELETE FROM dbo.TBUsuarios
            WHERE Cod_Usu = @Cod_Usu_norm;

            SELECT 'Usuario eliminado correctamente.' AS Mensaje, @Cod_Usu_norm AS Cod_Usu;
        END TRY
        BEGIN CATCH
            DECLARE @ErrMsg3 NVARCHAR(4000) = ERROR_MESSAGE();
            DECLARE @ErrNum3 INT = ERROR_NUMBER();
            DECLARE @ErrLine3 INT = ERROR_LINE();
            DECLARE @ErrProc3 SYSNAME = ERROR_PROCEDURE();
            DECLARE @ErrProcName3 SYSNAME = ISNULL(@ErrProc3, 'SP_TBUsuarios_CRUD');

            RAISERROR('Error %d en %s (línea %d): %s', 16, 1, @ErrNum3, @ErrProcName3, @ErrLine3, @ErrMsg3);
        END CATCH

        RETURN;
    END

    ----------------------------------------------------------------------
    -- SELECCIONAR (S)
    ----------------------------------------------------------------------
    IF @opcion = 'S'
    BEGIN
        IF @Cod_Usu_norm IS NULL
        BEGIN
            -- Orden numérico por IDUser si es posible; los no numéricos al final (NULL)
            SELECT *
            FROM dbo.TBUsuarios
            ORDER BY CASE WHEN IDUser NOT LIKE '%[^0-9]%' THEN CONVERT(INT, IDUser) ELSE NULL END;
        END
        ELSE
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.TBUsuarios WHERE Cod_Usu = @Cod_Usu_norm)
            BEGIN
                RAISERROR('No se encontró un usuario con el Cod_Usu indicado.', 16, 1);
                RETURN;
            END

            SELECT * FROM dbo.TBUsuarios WHERE Cod_Usu = @Cod_Usu_norm;
        END

        RETURN;
    END
END
GO









-- INSERT
EXEC dbo.SP_TBUsuarios_CRUD
     @opcion = 'I',
     @ApeNomUser = 'Prueba3',
     @Password = '23DE45D10BA40432C4E38F01F15C17';

-- SELECT TODOS
EXEC dbo.SP_TBUsuarios_CRUD @opcion = 'S';

-- UPDATE (usa el Cod_Usu que te devolvió el INSERT, ej. '0001')
EXEC dbo.SP_TBUsuarios_CRUD
     @opcion = 'U',
     @Cod_Usu = '0543',
     @ApeNomUser = 'EDIT',
     @Password = 'ABCDEF0123456789ABCDEF0123456789';

-- DELETE
EXEC dbo.SP_TBUsuarios_CRUD
     @opcion = 'D',
     @Cod_Usu = '0001';


















