Imports Treegram.ConstructionManagement

Public Module AecNaming
    Public ReadOnly DicoObjType As New Dictionary(Of String, String) From {{"IfcWall", "Murs"}, {"Murs", "IfcWall"},
                                                                  {"IfcSpace", "Pièces Archi"}, {"Pièces Archi", "IfcSpace"},
                                                                          {"IfcSlab", "Dalles"}, {"Dalles", "IfcSlab"},
                                                                          {"IfcStair", "Escaliers"}, {"Escaliers", "IfcStair"},
                                                                          {"IfcStairFlight", "Volées d'escaliers"}, {"Volées d'escaliers", "IfcStairFlight"},
                                                                          {"IfcColumn", "Poteaux"}, {"Poteaux", "IfcColumn"},
                                                                          {"IfcDoor", "Portes"}, {"Portes", "IfcDoor"}, {"Door", "Portes"},
                                                                          {"IfcWindow", "Fenêtres"}, {"Fenêtres", "IfcWindow"}, {"Window", "Fenêtres"},
                                                                          {"IfcOpeningElement", "Réservations"}, {"Réservations", "IfcOpeningElement"},
                                                                          {"IfcRailing", "Garde-corps"}, {"Garde-corps", "IfcRailing"},
                                                                          {"IfcBuildingElementPart", "Elements constructifs"}, {"Elements constructifs", "IfcBuildingElementPart"},
                                                                          {"IfcWallStandardCase", "Murs standards"}, {"Murs standards", "IfcWallStandardCase"},
                                                                          {"IfcBeam", "Poutres"}, {"Poutres", "IfcBeam"},
                                                                          {"IfcCovering", "Revêtement"}, {"Revêtement", "IfcCovering"},
                                                                          {"IfcCurtainWall", "Murs rideaux"}, {"CurtainWall", "Murs rideaux"},
                                                                          {"IfcRoof", "Toiture"}, {"Toiture", "IfcRoof"},
                                                                          {"IfcRamp", "Rampe"}, {"Rampe", "IfcRamp"},
                                                                          {"EnvelopeByTgm", "Enveloppes"}, {"Enveloppes", "EnvelopeByTgm"},
                                                                          {"ShellByTgm", "Coques"}, {"Coques", "ShellByTgm"},
                                                                          {"SpaceByTgm", "Pièces Tgm"}, {"Pièces Tgm", "SpaceByTgm"},
                                                                          {ProjectReference.Constants.overlappingObjType, "Eléments superposés"}, {"Eléments superposés", ProjectReference.Constants.overlappingObjType},
                                                                          {"ApartmentByTgm", "Appartements Tgm"}, {"Appartements Tgm", "ApartmentByTgm"},
                                                                          {"StairGroup", "Groupement d'escaliers"}, {"Groupement d'escaliers", "StairGroup"},
                                                                          {"ElevatorGroup", "Groupement d'ascenseurs"}, {"Groupement d'ascenseurs", "ElevatorGroup"},
                                                                          {"OverhangByTgm", "Surplombs"}, {"Surplombs", "OverhangByTgm"},
                                                                          {"TerraceByTgm", "Terrasses"}, {"Terrasses", "TerraceByTgm"},
                                                                          {"BalconyByTgm", "Balcons"}, {"Balcons", "BalconyByTgm"},
                                                                          {"FacadeByTgm", "Façades"}, {"Façades", "FacadeByTgm"},
                                                                          {"GlazingByTgm", "Surfaces vitrées"}, {"Surfaces vitrées", "GlazingByTgm"},
                                                                          {"IsVerticalSeparator", "Séparateurs verticaux"}, {"Séparateurs verticaux", "IsVerticalSeparator"}, 'Templates...
                                                                          {"IsHorizontalSeparator", "Séparateurs horizontaux"}, {"Séparateurs horizontaux", "IsHorizontalSeparator"}, 'Templates...                                                                      
                                                                          {"IsOpenable", "Ouvertures"}, {"Ouvertures", "IsOpenable"},
                                                                          {"IsSpace", "Pièces"}, {"Space", "Pièces"}, {"Pièces", "IsSpace"},
                                                                          {"IsSlab", "Planchers"}, {"Slab", "Planchers"}, {"Planchers", "IsSlab"},
                                                                          {"IsGlazed", "Eléments vitrés"}, {"Eléments vitrés", "IsGlazed"},
                                                                          {"WallOpeningEltByTgm", "Ouvertures dans murs"}, {"Ouvertures dans murs", "WallOpeningEltByTgm"}
                                                                          }
End Module
