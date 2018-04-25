# tom2amo
Library for converting between 1200 and 1103 Microsoft Analysis Services Tabular models

This library exists as a "sequel", so to speak, of the Tabular AMO 2012 library (found at https://github.com/juanpablojofre/tabularamo), which aimed to explain the labyrinthine and otherwise incomprehensible 1100 and 1103 Tabular Models.

This library aims to convert 1103, and possibly 1100, models to the 1200 and above TOM model (and vice-versa), as well as discard some of the "fluff" which exists purely to appease Visual Studio (for instance, the comments before measures). The reason for this, is that TOM has a one-to-one mapping of Tabular objects to TOM. As a result, it is far easier to manipulate.

In addition, the option to re-add compatibility with Visual Studio will be included.

Issues found in Tabular while creating this:

- In 1103, if a measure "A" and "A]" exist in the same model, and you add "A]" to a perspective, it will not add "A]", and will instead add "A". This issue does not exist in 1200.
- In 1103 and 1200, KPI's are created by creating hidden "measures". For instance, a KPI on the measure "A" creates the measures "_A Goal", "_A Status", and possibly "_A Trend". This means if you have a KPI "A", you cannot create a measure "_Goal".
- You cannot add KPIs to perspectives, only measures, in Visual Studio
