
Public Module Enumerations
    Public Enum constantGender
        M
        F
    End Enum

    Public Enum constantOwnerType
        SingleOwner
        JointOwner
    End Enum

    Public Enum SmokerStatus
        N           'No
        Y           'Yes
        C           'Cancer?????
    End Enum

    Public Enum PmtSolutionMethod
        SM          'Scalar Method
        IDM         'Initial Dollar Method
        ISM         'Inital Solution Method
        BSM         'Balloon Solutions Method
    End Enum

    Public Enum constantMortalityTableType
        SS           'Social Security 2009 Exact Age
        CDC          'CDC Life table for the total population: United States, 2009, SOURCE: CDC/NCHS, National Vital Statistics System.
        LSP          'Lifetime Security Plan data from HUD HECM
        VBT          '2001 Valuation Basic Table  -- Composite, Male, Female, Smoker, NonSmoker
        CSO          '2001 Valuation Basic Table  -- Composite, Male, Female, Smoker, NonSmoker 
        IAM          '2012 IAM Period Table
        ProIAM
        USLife
    End Enum

    Public Enum HomeTypes
        Condo
        Co_op
        DetachedSingle
        Duplex
        PUD
        ZeroLotLine
    End Enum

    Public Enum RoofTypes
        Asphalt
        BuiltUp
        CompositeShingle
        Metal
        SinglePlyBitumen
        Slate
        Tile
        WoodShake
    End Enum

    Public Enum ExteriorTypes
        AluminumVinyl
        Brick
        Stone
        Stucco
        Wood
    End Enum
    Public Enum PricingDefaultMultipliersX
        Type
        Locale
        AgeGrade
    End Enum

    Public Enum PricingDefaultPrevMaintsX
        PrevMaintAmt
        AppraisalCosts
        AppraisalUpdate
    End Enum

    Public Enum PricingDefaultRepairsX
        ExtPaint
        Roof
        Electrical
        Plumbing
        Structural
        Cleaning
        HVAC
        InteriorPaint
        FloorCoverings
        WindowCoverings
        Landscaping
        Damages
        Paving
    End Enum
End Module

