enum {		//{{AFX_DISP_ID(CMapView)
	
	// NOTE: ClassWizard will add and remove enumeration elements here
	//    DO NOT EDIT what you see in these blocks of generated code !
	// **ClassWizard is a thing of the past... feel free to edit this code.
	dispidSetSearchTolerance = 289L,
	dispidGetLayerFeatureByGeometry = 288L,
	dispidRemoveVolatilePoint = 287L,
	dispidAddVolatilePoint = 286L,
	dispidAddVolatileLayer = 285L,
	dispidQueryLayer = 284L,
	dispidAddDatasourceAndResave = 283L,
	dispidAddLayerAndResave = 282L,
	dispidGenerateLayerLabels = 281L,
	dispidSetLayerLabelFont = 280L,
	dispidSetLayerLabelShadow = 279L,
	dispidSetLayerLabelHalo = 278L,
	dispidSetLayerLabelScaling = 277L,
	dispidSetLayerLabelAttributes = 276L,
	dispidSetLayerFeatureColumn = 275L,
	dispidSetLayerLabelColumn = 274L,
	dispidSetLayerLabelMinScale = 273L,
	dispidSetLayerLabelMaxScale = 272L,
	dispidFreezeCurrentExtents = 271L,
	dispidSetConstrainingExtents = 270L,
	//^^^ Swift API ^^^
	dispidRecenterMapOnZoom = 262L,
	dispidUseAlternatePanCursor = 261L,
	dispidRedraw3 = 260L,
	dispidWmsLayer = 259L,
	dispidClearExtentHistory = 258L,
	dispidExtentHistoryRedoCount = 257L,
	dispidExtentHistoryUndoCount = 256L,
	dispidZoomToNext = 255L,
	dispidShowCoordinatesFormat = 254L,
	dispidLayerExtents = 253L,
	dispidCustomDrawingFlags = 252L,	
	dispidFocusRectangle = 251L,
	dispidIdentifiedShapes = 250L,
	
	dispidGeodesicArea = 249L,
	dispidGeodesicLength = 248L,
	dispidGeodesicDistance = 247L,
	dispidDrawLabelEx = 246L,
	dispidDrawLabel = 245L,
	dispidUndo = 244L,
	dispidIdentifier = 243L,
	dispidMouseTolerance = 242L,
	dispidUndoList = 241L,
	dispidLayerVisibleAtCurrentScale = 238L,
	dispidAddLayerFromDatabase = 237L,
	dispidOgrLayer = 236L,
	dispidShapeEditor = 235L,
	dispidZoomBarMaxZoom = 234L,
	dispidZoomBarMinZoom = 233L,
	dispidProjectionMismatchBehavior = 232L,
	dispidZoomBoxStyle = 231L,
	dispidZoomBarVerbosity = 230L,
	dispidReuseTileBuffer = 229L,
	dispidInertiaOnPanning = 228L,
	dispidAnimationOnZooming = 227L,
	dispidShowZoomBar = 226L,
	dispidDegreesToPixel = 225L,
	dispidPixelToDegrees = 224L,
	dispidDegreesToProj = 223L,
	dispidProjToDegrees = 222L,
	dispidGrabProjectionFromData = 221L,
	dispidRedraw2 = 220L,
	dispidShowCoordinates = 219L,
	dispidKnownExtents = 218L,
	dispidMapProjection = 217L,
	dispidLatitude = 212L,
	dispidLongitude = 213L,
	dispidCurrentZoom = 214L,
	dispidTileProvider = 216L,
	dispidFileManager = 211L,
	dispidAddLayerFromFilename = 210L,
	dispidGetKnownExtents = 209L,
	dispidSetGeographicExtents2 = 208L,
	dispidZoomBehavior = 207L,
	dispidClear = 206L,
	dispidScalebarUnits = 205L,
	dispidFindSnapPoint = 203L,
	dispidSnapShotToDC2 = 202L,
	dispidLoadTilesForSnapshot = 201L,
	dispidZoomToWorld = 200L,
	dispidLayerMinVisibleZoom = 199L,
	dispidLayerMaxVisibleZoom = 198L,
	dispidZoomToTileLevel = 197L,
	dispidMeasuring = 196L,
	dispidScalebarVisible = 195L,
	dispidSetGeographicExtents = 194L,
	dispidGeographicExtents = 193L,
	dispidProjection = 192L,
	dispidTiles = 191L,
	dispidZoomToSelected = 190L,
	dispidLayerFilename = 189,
	dispidPixelsPerDegree = 188,
	dispidMaxExtents = 187,
	dispidLayerSkipOnSaving = 186,
	dispidRemoveLayerOptions = 185L,
	dispidSerializeMapState = 184L,
	dispidDeserializeMapState = 183L,
	dispidLayerDescription = 182,
	dispidLoadLayerOptions = 181L,
	dispidSaveLayerOptions = 180L,
	dispidLoadMapState = 179L,
	dispidSaveMapState = 178L,
	dispidDeserializeLayerOptions = 177L,
	dispidSerializeLayerOptions = 176L,
	dispidImage = 175,
	dispidShapefile = 174,
	dispidShowVersionNumber = 173,
	dispidShowRedrawTime = 172,
	dispidLayerLabels = 171,
	dispidLayerDynamicVisibility = 169,
	dispidLayerMinVisibleScale = 168,
	dispidLayerMaxVisibleScale = 167,
	dispidVersionNumber = 166,
	dispidDrawWideCircleEx = 160L,
	dispidDrawWidePolygonEx = 161L,
	dispidBackColor = 1L,
	dispidZoomPercent = 2L,
	dispidCursorMode = 3L,
	dispidMapCursor = 4L,
	dispidUDCursorHandle = 5L,
	dispidSendMouseDown = 6L,
	dispidSendMouseUp = 7L,
	dispidSendMouseMove = 8L,
	dispidSendSelectBoxDrag = 9L,
	dispidSendSelectBoxFinal = 10L,
	dispidExtentPad = 11L,
	dispidExtentHistory = 12L,
	dispidKey = 13L,
	dispidDoubleBuffer = 14L,
	dispidGlobalCallback = 15L,
	dispidNumLayers = 16L,
	dispidExtents = 17L,
	dispidLastErrorCode = 18L,
	dispidIsLocked = 19L,
	dispidMapState = 20L,
	dispidInitializeMap = 21L,
	dispidUninitializeMap = 22L,
	dispidRedraw = 23L,
	dispidAddLayer = 24L,
	dispidRemoveLayer = 25L,
	dispidRemoveLayerWithoutClosing = 138L,
	dispidRemoveAllLayers = 26L,
	dispidMoveLayerUp = 27L,
	dispidMoveLayerDown = 28L,
	dispidMoveLayer = 29L,
	dispidMoveLayerTop = 30L,
	dispidMoveLayerBottom = 31L,
	dispidZoomToMaxExtents = 32L,
	dispidZoomToLayer = 33L,
	dispidZoomToShape = 34L,
	dispidZoomIn = 35L,
	dispidZoomOut = 36L,
	dispidZoomToPrev = 37L,
	dispidProjToPixel = 38L,
	dispidPixelToProj = 39L,
	dispidClearDrawing = 40L,
	dispidClearDrawings = 41L,
	dispidSnapShot = 42L,
	dispidApplyLegendColors = 43L,
	dispidLockWindow = 44L,
	dispidResize = 45L,
	dispidShowToolTip = 46L,
	dispidAddLabel = 47L,
	dispidClearLabels = 48L,
	dispidLayerFont = 49L,
	dispidGetColorScheme = 50L,
	dispidNewDrawing = 51L,
	dispidPoint = 52L,
	dispidLine = 53L,
	dispidCircle = 54L,
	dispidPolygon = 55L,
	dispidLayerKey = 56L,
	dispidLayerPosition = 57L,
	dispidLayerHandle = 58L,
	dispidLayerVisible = 59L,
	dispidShapeLayerFillColor = 60L,
	dispidShapeFillColor = 61L,
	dispidShapeLayerLineColor = 62L,
	dispidShapeLineColor = 63L,
	dispidShapeLayerPointColor = 64L,
	dispidShapePointColor = 65L,
	dispidShapeLayerDrawFill = 66L,
	dispidShapeDrawFill = 67L,
	dispidShapeLayerDrawLine = 68L,
	dispidShapeDrawLine = 69L,
	dispidShapeLayerDrawPoint = 70L,
	dispidShapeDrawPoint = 71L,
	dispidShapeLayerLineWidth = 72L,
	dispidShapeLineWidth = 73L,
	dispidShapeLayerPointSize = 74L,
	dispidShapePointSize = 75L,
	dispidShapeLayerFillTransparency = 76L,
	dispidShapeFillTransparency = 77L,
	dispidShapeLayerLineStipple = 78L,
	dispidShapeLineStipple = 79L,
	dispidShapeLayerFillStipple = 80L,
	dispidShapeFillStipple = 81L,
	dispidShapeVisible = 82L,
	dispidImageLayerPercentTransparent = 83L,
	dispidErrorMsg = 84L,
	dispidDrawingKey = 85L,
	dispidShapeLayerPointType = 86L,
	dispidShapePointType = 87L,
	dispidLayerLabelsVisible = 88L,
	dispidUDLineStipple = 89L,
	dispidUDFillStipple = 90L,
	dispidUDPointType = 91L,
	dispidGetObject = 92L,
	dispidLayerName = 91,	
	dispidSetImageLayerColorScheme = 92L,		
	dispidGridFileName = 93,	
	dispidUpdateImage = 94L,
	dispidSerialNumber = 95,
	dispidLineSeparationFactor = 96,
	dispidLayerLabelsShadow = 97L,
	dispidLayerLabelsScale = 98L,
	dispidAddLabelEx = 99L,
	dispidGetLayerStandardViewWidth = 100L,
	dispidSetLayerStandardViewWidth = 101L,
	dispidLayerLabelsOffset = 102L,
	dispidLayerLabelsShadowColor = 103L,
	dispidUseLabelCollision = 104L,
	dispidIsTIFFGrid = 105L,
	dispidIsSameProjection = 106L,
	dispidZoomToMaxVisibleExtents = 107L,
	dispidMapResizeBehavior = 108L,
	dispidDrawLineEx = 115L,		
	dispidDrawPointEx = 116L,		
	dispidDrawCircleEx = 117L,		
	dispidSendOnDrawBackBuffer = 118L,		
	dispidLabelColor = 119L,		
	dispidSetDrawingLayerVisible = 120L,		
	dispidClearDrawingLabels = 121L,
	dispidDrawingFont = 122L,
	dispidAddDrawingLabelEx = 123L,
	dispidAddDrawingLabel =124L,
	dispidDrawingLabelsOffset = 125L,
	dispidDrawingLabelsScale = 126L,
	dispidDrawingLabelsShadow = 127L,
	dispidDrawingLabelsShadowColor = 128L,
	dispidUseDrawingLabelCollision = 129L,
	dispidDrawingLabelsVisible = 130L,
	dispidGetDrawingStandardViewWidth = 131L,
	dispidSetDrawingStandardViewWidth = 132L,
	dispidMultilineLabels = 133,
	dispidSnapShot2 = 136L,
	dispidLayerFontEx = 137L, 
	dispidset_UDPointFontCharFont = 139L, 
	dispidset_UDPointFontCharFontSize = 141L,
	dispidset_UDPointFontCharListAdd = 140L, 
	dispidShapePointFontCharListID = 142L,
    dispidReSourceLayer = 143L,
	dispidShapeLayerStippleColor = 144L,
	dispidShapeStippleColor = 145L,
	dispidShapeStippleTransparent = 146L,
	dispidShapeLayerStippleTransparent = 147L,
    dispidTrapRMouseDown = 148L,
	dispidDisableWaitCursor = 149L,
	dispidAdjustLayerExtents = 150L,
	dispidUseSeamlessPan = 151L,
	dispidMouseWheelSpeed = 152L,
	dispidSnapShot3 = 153L,
	dispidShapeDrawingMethod = 154L,
	dispidDrawPolygonEx = 155L,
	dispidCurrentScale = 156L,
	dispidDrawingLabels = 157L,
	dispidMapUnits = 158L,
	dispidSnapShotToDC = 159L,
    dispidMapRotationAngle = 162L,           //ajp June 2010
    dispidRotatedExtent = 163L,				  //ajp June 2010
    dispidGetBaseProjectionPoint = 164L,     //ajp June 2010
	dispidCanUseImageGrouping = 165L,		 // lsu
	dispidDrawBackBuffer = 170L,

	// events
	eventidMouseDown = 1L,
	eventidMouseUp = 2L,
	eventidMouseMove = 3L,
	eventidFileDropped = 4L,
	eventidSelectBoxFinal = 5L,
	eventidSelectBoxDrag = 6L,
	eventidExtentsChanged = 7L,
	eventidMapState = 8L,
	eventidOnDrawBackBuffer = 9L,
	eventidShapeHighlighted = 10L,
	eventidBeforeDrawing = 11L,
	eventidAfterDrawing = 12L,
	eventidTilesLoaded = 13L,
	eventidMeasuringChanged = 14L,
	eventidBeforeShapeEdit = 16L,
	eventidValidateShape = 17L,
	eventidAfterShapeEdit = 18L,
	eventidChooseLayer = 19L,
	eventidShapeValidationFailed = 21L,
	eventidBeforeDeleteShape = 22L,
	eventidProjectionChanged = 23L,
	eventidUndoListChanged = 24L,
	eventidSelectionChanged = 25L,
	eventidShapeIdentified = 26L,
	eventidLayerProjectionIsEmpty = 27L,
	eventidProjectionMismatch = 28L,
	eventidLayerReprojected = 29L,
	eventidLayerAdded = 30L,
	eventidLayerRemoved = 31L,
	eventidBackgroundLoadingStarted = 32L,
	eventidBackgroundLoadingFinished = 33L,
	eventidGridOpened = 34L,
	eventidShapesIdentified = 35L,
	eventidOnDrawBackBuffer2 = 36L,
	eventidBeforeLayers = 37L,
	eventidAfterLayers = 38L,
	//}}AFX_DISP_ID
};
