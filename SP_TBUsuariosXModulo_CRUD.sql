USE BDThimble_QA;
GO

IF EXISTS (SELECT 1 FROM sys.objects 
           WHERE name = 'SP_TBUsuariosXModulo_CRUD' AND type = 'P')
    DROP PROCEDURE SP_TBUsuariosXModulo_CRUD;
GO

CREATE PROCEDURE SP_TBUsuariosXModulo_CRUD
(
    @opcion CHAR(1),         -- I / D / S
    @IDUser CHAR(8)             
)
AS
BEGIN
    SET NOCOUNT ON;

    ---------------------------------------------------------------------
    -- Lista fija de módulos a asignar
    ---------------------------------------------------------------------
    DECLARE @Modulos TABLE (IDModulo CHAR(8));
    INSERT INTO @Modulos(IDModulo)
    VALUES ('04010001'), ('04010002'), ('04010003'), ('04010004');

    ---------------------------------------------------------------------
    -- Validación inicial
    ---------------------------------------------------------------------
    IF (@IDUser IS NULL OR LTRIM(RTRIM(@IDUser)) = '')
    BEGIN
        RAISERROR('El parámetro @IDUser es obligatorio.', 16, 1);
        RETURN;
    END

    ---------------------------------------------------------------------
    -- OPCIÓN: SELECT
    -- Muestra solo el IDUser existente en TBUsuariosXModulo.
    ---------------------------------------------------------------------
    IF (@opcion = 'S')
    BEGIN
        SELECT DISTINCT IDUser
        FROM TBUsuariosXModulo
        WHERE IDUser = @IDUser;

        RETURN;
    END

    ---------------------------------------------------------------------
    -- OPCIÓN: DELETE
    -- Elimina todos los módulos asignados a un usuario.
    ---------------------------------------------------------------------
    IF (@opcion = 'D')
    BEGIN
        IF NOT EXISTS (SELECT 1 FROM TBUsuariosXModulo WHERE IDUser = @IDUser)
        BEGIN
            RAISERROR('No existen registros para el usuario indicado.', 16, 1);
            RETURN;
        END

        DELETE FROM TBUsuariosXModulo WHERE IDUser = @IDUser;

        RETURN;
    END

    ---------------------------------------------------------------------
    -- OPCIÓN: INSERT
    -- Inserta los 4 módulos por cada usuario recién creado.
    ---------------------------------------------------------------------
    IF (@opcion = 'I')
    BEGIN
        -----------------------------------------------------------------
        -- Verificar que usuario no tenga módulos ya generados
        -----------------------------------------------------------------
        IF EXISTS (SELECT 1 FROM TBUsuariosXModulo WHERE IDUser = @IDUser)
        BEGIN
            RAISERROR('El usuario ya tiene módulos asignados.', 16, 1);
            RETURN;
        END

        -----------------------------------------------------------------
        -- Obtener siguiente correlativo IDUsuarioModulo (char(8))
        -----------------------------------------------------------------
        DECLARE @UltimoID CHAR(8) =
            (SELECT TOP 1 IDUsuarioModulo
             FROM TBUsuariosXModulo
             ORDER BY IDUsuarioModulo DESC);

        DECLARE @NextID INT;

        IF @UltimoID IS NULL
            SET @NextID = 1;
        ELSE
            SET @NextID = CAST(@UltimoID AS INT) + 1;

        -----------------------------------------------------------------
        -- Insertar los 4 módulos
        -----------------------------------------------------------------
        INSERT INTO TBUsuariosXModulo
        (
            IDUsuarioModulo,
            IDUser,
            IDModulo,
            FecIni,
            FecFin,
            Nivel,
            Estado
        )
        SELECT
            RIGHT('00000000' + CAST(ROW_NUMBER() OVER (ORDER BY IDModulo) + @NextID - 1 AS VARCHAR(8)), 8),
            @IDUser,
            M.IDModulo,
            CONVERT(CHAR(8), GETDATE(), 112),    -- Hoy (AAAAMMDD)
            NULL,                                
            '00',                                -- Nivel estándar
            'A'                                  -- Activo
        FROM @Modulos M;

        RETURN;
    END

    ---------------------------------------------------------------------
    -- Opción inválida
    ---------------------------------------------------------------------
    RAISERROR('La opción indicada no es válida. Use I, D o S.', 16, 1);
END
GO



EXEC SP_TBUsuariosXModulo_CRUD 'I', '00001234';
go
EXEC SP_TBUsuariosXModulo_CRUD 'S', '00001234';
go
select*from TBUsuariosXModulo
go



EXEC SP_TBUsuariosXModulo_CRUD 'D', '00001234';
go

EXEC SP_TBUsuariosXModulo_CRUD 'S', '00001234';
go
select*from TBUsuariosXModulo
go


select*from TBUsuariosXModulo
go