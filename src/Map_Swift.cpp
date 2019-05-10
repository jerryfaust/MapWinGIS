// Swift-specific methods added for custom behavior

#pragma once
#include "stdafx.h"
#include "Map.h"
#include "ComHelpers\ProjectionHelper.h"

enum volatileLayerType
{
	vltPolygon = 1,
	vltPolyline,
	vltPoint
};

// volatile layer references
long _PointLayerHandle = -1;
long _PolylineLayerHandle = -1;
long _PolygonLayerHandle = -1;
CComPtr<IShapefile> _PointLayer = NULL;
CComPtr<IShapefile> _PolylineLayer = NULL;
CComPtr<IShapefile> _PolygonLayer = NULL;

// reprojector, from WGS84 to the current map projection
CComPtr<IGeoProjection> _WGS84 = NULL;

double SearchTolerance = 10.0;

VARIANT_BOOL CMapView::SetConstrainingExtents(DOUBLE xMin, DOUBLE yMin, DOUBLE xMax, DOUBLE yMax)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// save parameters
	_xMin = xMin;
	_yMin = yMin;
	_xMax = xMax;
	_yMax = yMax;
	// turn it on
	_useConstrainingExtents = TRUE;

	return VARIANT_TRUE;
}

void CMapView::FreezeCurrentExtents()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// save current extents
	_xMin = _extents.left;
	_yMin = _extents.bottom;
	_xMax = _extents.right;
	_yMax = _extents.top;
	// turn it on
	_useConstrainingExtents = TRUE;
}

// ****************************************************************** 
//		SetLayerLabelMaxScale
// ****************************************************************** 
void CMapView::SetLayerLabelMaxScale(LONG LayerHandle, DOUBLE maxScale)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Layer* layer = GetLayer(LayerHandle);
	if (layer)
	{
		ILabels* lbls = layer->get_Labels();
		if (lbls)
		{
			lbls->put_DynamicVisibility(VARIANT_TRUE);
			lbls->put_MaxVisibleScale(maxScale);
		}
	}
}

// ****************************************************************** 
//		SetLayerLabelMinScale
// ****************************************************************** 
void CMapView::SetLayerLabelMinScale(LONG LayerHandle, DOUBLE MinScale)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	Layer* layer = GetLayer(LayerHandle);
	if (layer)
	{
		ILabels* lbls = layer->get_Labels();
		if (lbls)
		{
			lbls->put_DynamicVisibility(VARIANT_TRUE);
			lbls->put_MinVisibleScale(MinScale);
		}
	}
}


// ****************************************************************** 
//		LayerMinVisibleZoom
// ****************************************************************** 
int CMapView::GetLayerMinVisibleZoom(LONG LayerHandle)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	Layer* layer = GetLayer(LayerHandle);
	return layer ? layer->minVisibleZoom : -1;
}

void CMapView::SetLayerMinVisibleZoom(LONG LayerHandle, int newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	Layer* layer = GetLayer(LayerHandle);
	if (layer) {
		if (newVal < 0) newVal = 0;
		if (newVal > 18) newVal = 18;
		layer->minVisibleZoom = newVal;
	}
}

// ****************************************************************** 
//		LayerMaxVisibleZoom
// ****************************************************************** 
int CMapView::GetLayerMaxVisibleZoom(LONG LayerHandle)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Layer* layer = GetLayer(LayerHandle);
	return layer ? layer->maxVisibleZoom : -1;
}

void CMapView::SetLayerMaxVisibleZoom(LONG LayerHandle, int newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	Layer* layer = GetLayer(LayerHandle);
	if (layer)
	{
		if (newVal < 0) newVal = 0;
		if (newVal > 100) newVal = 100;
		layer->maxVisibleZoom = newVal;
	}
}

// ****************************************************************** 
//		SetLayerFeatureColumn
// ****************************************************************** 
void CMapView::SetLayerFeatureColumn(LONG LayerHandle, LPCTSTR ColumnName)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// get Shapefile reference
	CComPtr<IShapefile> sf = this->GetShapefile(LayerHandle);
	if (sf)
	{
		long fieldIndex = -1;
		// save feature columns (unpack into field indices
		CAtlString strColumns(ColumnName);
		// is it a single or multi-field label?
		if (strColumns.Find("\x1F") < 0)
		{
			// single-field labeling
			CComBSTR bstrColName(strColumns);
			sf->get_FieldIndexByName(bstrColName.Copy(), &fieldIndex);
			_featureColumns[LayerHandle].push_back(fieldIndex);
		}
		else
		{
			// multi-field labeling
			int tokenPos = 0;
			CAtlString strToken = strColumns.Tokenize("\x1F", tokenPos);
			while (!strToken.IsEmpty())
			{
				CComBSTR bstrColName(strToken);
				sf->get_FieldIndexByName(bstrColName.Copy(), &fieldIndex);
				_featureColumns[LayerHandle].push_back(fieldIndex);
				// next?
				strToken = strColumns.Tokenize("\x1F", tokenPos);
			}
		}
	}
}


// ****************************************************************** 
//		SetLayerLabelColumn
// ****************************************************************** 
void CMapView::SetLayerLabelColumn(LONG LayerHandle, LPCTSTR ColumnName)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// save label columns
	_labelColumns[LayerHandle] = CAtlString(ColumnName);
	// pass on to SetLayerLabelAttributes
	SetLayerLabelAttributes(LayerHandle, "Arial", 10, FALSE);
}


// ****************************************************************** 
//		SetLayerLabelAttributes
// ****************************************************************** 
void CMapView::SetLayerLabelAttributes(LONG LayerHandle, LPCTSTR FontName, LONG FontSize, BOOL ScaleFonts)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// make sure we have the Label column name
	CAtlString strColumn = _labelColumns[LayerHandle];
	if (strColumn.GetLength() > 0)
	{
		// set up Labeling
		IShapefile* sf = GetShapefile(LayerHandle);
		if (sf)
		{
			// Labels reference
			ILabels* labels = nullptr;
			sf->get_Labels(&labels);
			if (labels != nullptr)
			{
				// general label settings
				CComBSTR bstr(FontName);
				labels->put_UseGdiPlus(VARIANT_TRUE);
				labels->put_RemoveDuplicates(VARIANT_TRUE);
				labels->put_Visible(VARIANT_TRUE);
				labels->put_FrameVisible(VARIANT_FALSE);
				labels->put_FontName(bstr.Copy());
				labels->put_FontSize(FontSize);
				//labels->put_HaloVisible(VARIANT_TRUE);
				//labels->put_HaloSize(5);
				//labels->put_FontOutlineVisible(VARIANT_TRUE);
				//labels->put_FontOutlineColor(0xFFFFFF);
				labels->put_AvoidCollisions(VARIANT_TRUE);
				labels->put_ScaleLabels(ScaleFonts);
				if (ScaleFonts)
				{
					labels->put_FontSize2(FontSize + 2);
				}

				//// get layer drawing options
				//IShapeDrawingOptions* options = nullptr;
				//sf->get_DefaultDrawingOptions(&options);
				//if (options != nullptr)
				//{
				//// label positioning
				//tkLabelPositioning labelPos = tkLabelPositioning::lpNone;
				//BOOL largestPartOnly = FALSE;

				// what is the basic geometry type
				ShpfileType sfType;
				sf->get_ShapefileType2D(&sfType);
				if (sfType == ShpfileType::SHP_POINT)
				{
					//// center label below
					//labelPos = tkLabelPositioning::lpCenter;
					// use square frame
					labels->put_FrameVisible(VARIANT_TRUE);
					labels->put_FrameType(tkLabelFrameType::lfRectangle);
					labels->put_FrameOutlineColor(0);
					labels->put_InboxAlignment(tkLabelAlignment::laBottomCenter);
					// labels
					labels->put_Alignment(tkLabelAlignment::laCenter);
					labels->put_AutoOffset(VARIANT_TRUE);
				}
				else if (sfType == ShpfileType::SHP_POLYLINE)
				{
					//// center label above longest segment
					//largestPartOnly = TRUE;
					//labelPos = tkLabelPositioning::lpLongestSegement;
					// no frame
					labels->put_FrameVisible(VARIANT_FALSE);
					// labels
					labels->put_Alignment(tkLabelAlignment::laTopCenter);
					labels->put_AutoOffset(VARIANT_TRUE);
				}
				else if (sfType == ShpfileType::SHP_POLYGON)
				{
					//// center label
					//labelPos = tkLabelPositioning::lpCentroid;
					// no frame ?
					labels->put_FrameVisible(VARIANT_FALSE);
				}
				//long fieldIndex = -1;
				//long count = 0;
				//CComBSTR bstrColName(strColumn);
				//sf->get_FieldIndexByName(bstrColName.Copy(), &fieldIndex);
				//sf->GenerateLabels(fieldIndex, labelPos, largestPartOnly, &count);

				//	// release the reference
				//	options->Release();
				//}
				// release Labels reference
				labels->Release();
			}
		}
	}
}


// ****************************************************************** 
//		SetLayerLabelScaling
// ****************************************************************** 
void CMapView::SetLayerLabelScaling(LONG LayerHandle, LONG FontSize, DOUBLE RelativeScale)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// get Shapefile reference
	IShapefile* sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// Labels reference
		ILabels* labels = nullptr;
		sf->get_Labels(&labels);
		if (labels != nullptr)
		{
			//set up Font scaling
			labels->put_ScaleLabels(VARIANT_TRUE);
			labels->put_BasicScale(RelativeScale);
			labels->put_FontSize(FontSize);

			// release Labels reference
			labels->Release();
		}
	}
}


// ****************************************************************** 
//		SetLayerLabelHalo
// ****************************************************************** 
void CMapView::SetLayerLabelHalo(LONG LayerHandle, LONG HaloSize, LONG HaloColor)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// get Shapefile reference
	IShapefile* sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// Labels reference
		ILabels* labels = nullptr;
		sf->get_Labels(&labels);
		if (labels != nullptr)
		{
			labels->put_HaloVisible(VARIANT_TRUE);
			labels->put_HaloSize(HaloSize);
			labels->put_HaloColor(HaloColor);

			// release Labels reference
			labels->Release();
		}
	}
}


// ****************************************************************** 
//		SetLayerLabelShadow
// ****************************************************************** 
void CMapView::SetLayerLabelShadow(LONG LayerHandle, LONG OffsetX, LONG OffsetY, LONG ShadowColor)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// get Shapefile reference
	IShapefile* sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// Labels reference
		ILabels* labels = nullptr;
		sf->get_Labels(&labels);
		if (labels != nullptr)
		{
			labels->put_ShadowVisible(VARIANT_TRUE);
			labels->put_ShadowOffsetX(OffsetX);
			labels->put_ShadowOffsetY(OffsetY);
			labels->put_ShadowColor(ShadowColor);

			// release Labels reference
			labels->Release();
		}
	}
}


// ****************************************************************** 
//		SetLayerLabelFont
// ****************************************************************** 
void CMapView::SetLayerLabelFont(LONG LayerHandle, LPCTSTR FontName, LONG FontSize, LONG FontColor, BOOL FontBold, BOOL FontItalic, DOUBLE RelativeScale)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// get Shapefile reference
	IShapefile* sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// Labels reference
		ILabels* labels = nullptr;
		sf->get_Labels(&labels);
		if (labels != nullptr)
		{
			// set Font attributes
			CComBSTR bstr(FontName);
			labels->put_FontName(bstr);
			labels->put_FontSize(FontSize);
			labels->put_FontColor(FontColor);
			labels->put_FontBold(FontBold ? VARIANT_TRUE : VARIANT_FALSE);
			labels->put_FontItalic(FontItalic ? VARIANT_TRUE : VARIANT_FALSE);

			// is scaling requested?
			if (RelativeScale != 0.0)
			{
				//set up Font scaling
				labels->put_ScaleLabels(VARIANT_TRUE);
				labels->put_BasicScale(RelativeScale);
			}

			// release Labels reference
			labels->Release();
		}
	}
}

// local helper function
CAtlString BuildLabelExpression(CAtlString input)
{
	CAtlString exp;
	int tokenPos = 0;
	CAtlString strToken = input.Tokenize("\037", tokenPos);
	while (!strToken.IsEmpty())
	{
		// watch for secondary '+' separator, which indicates
		// concatenation with no space between

		if (strToken.Find('+') >= 0)
		{
			// if expression is in progress, we need to add trailing space
			if (exp.GetLength() > 0)
				// add trailing space for next major token
				exp.Format("%s + \" \"", exp);

			int plusPos = 0;
			// now process plus-separated tokens
			CAtlString strPlus = strToken.Tokenize("+", plusPos);
			while (!strPlus.IsEmpty())
			{
				// if empty, add first field value
				if (exp.GetLength() == 0)
					exp.Format("[%s]", strPlus);
				else
					// plus sign indicates concatenation (w/ no spaces)
					exp.Format("%s + [%s]", exp, strPlus);
				// next?
				strPlus = strToken.Tokenize("+", plusPos);
			}
		}
		else
		{
			// else add field values with space between

			// if empty, add first field value
			if (exp.GetLength() == 0)
				exp.Format("[%s]", strToken);
			else
				// add new field value with space between
				exp.Format("%s + \" \" + [%s]", exp, strToken);
		}
		// next?
		strToken = input.Tokenize("\037", tokenPos);
	}
	// and return
	return exp;
}

// ****************************************************************** 
//		GenerateLabels
// ****************************************************************** 
void CMapView::GenerateLayerLabels(LONG LayerHandle)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// get Shapefile reference
	IShapefile* sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// Labels reference
		ILabels* labels = nullptr;
		sf->get_Labels(&labels);
		if (labels != nullptr)
		{
			// label positioning
			tkLabelPositioning labelPos = tkLabelPositioning::lpNone;
			BOOL largestPartOnly = FALSE;
			// what is the basic geometry type
			ShpfileType sfType;
			sf->get_ShapefileType2D(&sfType);
			if (sfType == ShpfileType::SHP_POINT)
			{
				// center label below
				labelPos = tkLabelPositioning::lpCenter;
			}
			else if (sfType == ShpfileType::SHP_POLYLINE)
			{
				// center label above longest segment
				largestPartOnly = TRUE;
				labelPos = tkLabelPositioning::lpLongestSegement;
			}
			else if (sfType == ShpfileType::SHP_POLYGON)
			{
				// center label
				labelPos = tkLabelPositioning::lpCentroid;
			}

			// re-generate labels
			long fieldIndex = -1;
			long count = 0;
			CAtlString strColumn = _labelColumns[LayerHandle];
			// is it a single or multi-field label?
			if (strColumn.Find('\037') < 0)
			{
				// single-field labeling
				CComBSTR bstrColName(strColumn);
				sf->get_FieldIndexByName(bstrColName.Copy(), &fieldIndex);
				sf->GenerateLabels(fieldIndex, labelPos, largestPartOnly, &count);
			}
			else
			{
				// multi-field labeling
				CAtlString strExpr = BuildLabelExpression(strColumn);
				CComBSTR bstrExpr(strExpr);
				labels->Generate(bstrExpr, labelPos, largestPartOnly, &count);
			}

			// release Labels reference
			labels->Release();
		}
	}
}

// ***************************************************************
//		AddLayerAndResave()
// ***************************************************************
long CMapView::AddLayerAndResave(LPCTSTR Filename, BOOL visible)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	USES_CONVERSION;
	IDispatch* layer = NULL;
	CComBSTR bstr(Filename);
	// open the file
	_fileManager->Open(bstr, tkFileOpenStrategy::fosAutoDetect, _globalCallback, &layer);
	if (layer)
	{
		long handle = -1;
		// get Shapefile interface
		CComPtr<IShapefile> sf = NULL;
		layer->QueryInterface(IID_IShapefile, (void**)&sf);
		// is it a Shapefile ?
		if (sf != NULL)
		{
			// see if projections differ before adding layer, since we would otherwise have to
			// unload the layer to unlock it, save over it, then load it again...

			// get the projections
			CComPtr<IGeoProjection> gpMap = NULL;
			CComPtr<IGeoProjection> gpLayer = NULL;
			gpMap = this->GetGeoProjection();
			sf->get_GeoProjection(&gpLayer);
			// map may not have a projection yet...
			VARIANT_BOOL isEmpty;
			if (gpMap && (gpMap->get_IsEmpty(&isEmpty) == S_OK) && (isEmpty == VARIANT_FALSE))
			{
				// mismatch testing
				Layer* lyr = new Layer();
				lyr->set_Object(sf);
				CComPtr<IExtents> bounds = NULL;
				lyr->GetExtentsAsNewInstance(&bounds);
				// do they differ?
				if (!ProjectionHelper::IsSame(gpMap, gpLayer, bounds, 20))
				{
					// rename exising shapefile
					CAtlString sfName(Filename);
					CAtlString newName(sfName);
					newName.MakeLower().Replace(".shp", ".original.shp");
					CopyFile((LPCTSTR)sfName, (LPCTSTR)newName, TRUE);
					// rename exising shx
					sfName.Replace(".shp", ".shx");
					newName = sfName;
					newName.Replace(".shx", ".original.shx");
					CopyFile((LPCTSTR)sfName, (LPCTSTR)newName, TRUE);
					// rename exising dbf
					sfName.Replace(".shx", ".dbf");
					newName = sfName;
					newName.Replace(".dbf", ".original.dbf");
					CopyFile((LPCTSTR)sfName, (LPCTSTR)newName, TRUE);
					// rename exising prj
					sfName.Replace(".dbf", ".prj");
					newName = sfName;
					newName.Replace(".prj", ".original.prj");
					CopyFile((LPCTSTR)sfName, (LPCTSTR)newName, TRUE);
					Sleep(100);

					// reproject to a new in-memory Shapefile
					CComPtr<IShapefile> sfNew = NULL;
					long count;
					sf->Reproject(gpMap, &count, &sfNew);

					// close and delete the originals
					VARIANT_BOOL vb;
					sf->Close(&vb);
					layer->Release();
					sfName = Filename;
					sfName.MakeLower();
					remove((LPCTSTR)sfName);
					sfName.Replace(".shp", ".shx");
					remove((LPCTSTR)sfName);
					sfName.Replace(".shx", ".dbf");
					remove((LPCTSTR)sfName);
					sfName.Replace(".dbf", ".prj");
					remove((LPCTSTR)sfName);
					Sleep(100);

					// now resave, to the original name, in the new projection
					VARIANT_BOOL success;
					sfNew->SaveAs(bstr, _globalCallback, &success);

					bstr = Filename;
					// open the reprojected file
					_fileManager->Open(bstr, tkFileOpenStrategy::fosAutoDetect, _globalCallback, &layer);
					if (layer == NULL) return -1;
				}
			}
		}
		// add to the map
		handle = AddLayer(layer, visible);

		// let it go
		layer->Release();
		return handle;
	}
	else
	{
		// failed
		return -1;
	}
}

// ***************************************************************
//		AddDatasourceAndResave()
// ***************************************************************
long CMapView::AddDatasourceAndResave(LPCTSTR ConnectionString, LPCTSTR TableName, BOOL visible)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	USES_CONVERSION;
	long handle = -1;
	IOgrLayer* layer = NULL;
	CComBSTR bstrConnection(ConnectionString);
	CComBSTR bstrTable(TableName);
	// build Shapefile path

	// if placing in parent directory
	CAtlString strDirectory(ConnectionString);
	int pos = strDirectory.ReverseFind('\\');
	strDirectory = strDirectory.Left(pos);
	CComBSTR bstrShapefile((LPCTSTR)strDirectory);

	// if placing in GDB directory
	//CComBSTR bstrShapefile(ConnectionString);

	bstrShapefile.Append("\\");
	bstrShapefile.Append(bstrTable);
	bstrShapefile.Append(".shp");

	// does the Shapefile already exist?
	handle = AddLayerFromFilename(CAtlString(bstrShapefile), tkFileOpenStrategy::fosAutoDetect, visible);
	if (handle >= 0) return handle;

	// else continue on loading from database
	_fileManager->OpenFromDatabase(bstrConnection, bstrTable, &layer);
	if (layer)
	{
		VARIANT_BOOL vbResult;
		layer->get_SupportsEditing(tkOgrSaveType::ostSaveAll, &vbResult);
		// for now, only if a GDB datasource
		CAtlString strConnStr(ConnectionString);
		if (strConnStr.MakeLower().Find(".gdb") > 0)
		{
			// see if projections differ before adding layer, since we would otherwise have to
			// unload the layer to unlock it, save over it, then load it again...

			// get the projections
			CComPtr<IGeoProjection> gpMap = NULL;
			CComPtr<IGeoProjection> gpLayer = NULL;
			gpMap = this->GetGeoProjection();
			layer->get_GeoProjection(&gpLayer);
			// map may not have a projection yet...
			VARIANT_BOOL isEmpty;
			if (gpMap && (gpMap->get_IsEmpty(&isEmpty) == S_OK) && (isEmpty == VARIANT_FALSE))
			{
				// mismatch testing
				Layer* lyr = new Layer();
				lyr->set_Object(layer);
				CComPtr<IExtents> bounds = NULL;
				lyr->GetExtentsAsNewInstance(&bounds);
				// do they differ?
				if (!ProjectionHelper::IsSame(gpMap, gpLayer, bounds, 20))
				{
					// get OGR layer as Shapefile
					CComPtr<IShapefile> sf = NULL;
					layer->GetBuffer(&sf);
					// reproject to a new in-memory Shapefile
					CComPtr<IShapefile> sfNew = NULL;
					long count;
					sf->Reproject(gpMap, &count, &sfNew);
					// close original
					VARIANT_BOOL vb;
					sf->Close(&vb);
					layer->Release();
					// now save, to Shapefile of original name, in the new projection
					VARIANT_BOOL success;
					sfNew->SaveAs(bstrShapefile, _globalCallback, &success);
					// add to map
					handle = AddLayer(sfNew, visible);
					return handle;
				}
			}
		}
		handle = AddLayer(layer, visible);
		layer->Release();
		return handle;
	}
	else
	{
		return -1;
	}
}

// ***************************************************************
//		QueryLayer()
// ***************************************************************
BSTR CMapView::QueryLayer(LONG LayerHandle, LPCTSTR WhereClause)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CAtlString strResult("Query Results");
	CComBSTR result("Query Results");

	IShapefile* sf = GetShapefile(LayerHandle);
	if (sf)
	{
		ITable* tab = nullptr;
		sf->get_Table(&tab);
		if (tab)
		{
			CComBSTR expression(WhereClause);
			CComBSTR errorString;
			VARIANT results;
			VARIANT_BOOL vb;
			tab->Query(expression, &results, &errorString, &vb);
			if (vb != VARIANT_FALSE)
			{

			}
		}
	}

	//return result;
	return A2BSTR((LPCTSTR)strResult);
}

// ***************************************************************
//		SetSearchTolerance()
// ***************************************************************
void CMapView::SetSearchTolerance(DOUBLE Tolerance)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	//
	SearchTolerance = Tolerance;
}

// ***************************************************************
//		AddVolatileLayer()
// ***************************************************************
LONG CMapView::AddVolatileLayer(LONG GeometryType, BOOL Visible)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	VARIANT_BOOL vb;
	LONG layerHandle = -1;

	// set up layer
	switch ((volatileLayerType)GeometryType)
	{
	case vltPoint:
		ComHelper::CreateInstance(idShapefile, (IDispatch**)&_PointLayer);
		_PointLayer->CreateNew(CComBSTR(""), ShpfileType::SHP_POINT, &vb);
		if (vb == VARIANT_TRUE)
		{
			_PointLayer->put_Volatile(VARIANT_TRUE);
			_PointLayer->put_Selectable(VARIANT_TRUE);
			// add layer to map
			layerHandle = _PointLayerHandle = this->AddLayer((LPDISPATCH)_PointLayer, Visible);
			// rendering?
			this->SetShapeLayerPointColor(_PointLayerHandle, 255);
			this->SetShapeLayerPointSize(_PointLayerHandle, 14);
			this->SetShapeLayerPointType(_PointLayerHandle, 2); // star
		}
		break;
	case vltPolyline:
		ComHelper::CreateInstance(idShapefile, (IDispatch**)&_PolylineLayer);
		_PolylineLayer->CreateNew(CComBSTR(""), ShpfileType::SHP_POLYLINE, &vb);
		if (vb == VARIANT_TRUE)
		{
			_PolylineLayer->put_Volatile(VARIANT_TRUE);
			_PolylineLayer->put_Selectable(VARIANT_TRUE);
			// add layer to map
			layerHandle = _PolylineLayerHandle = this->AddLayer((LPDISPATCH)_PolylineLayer, Visible);
		}
		break;
	case vltPolygon:
		ComHelper::CreateInstance(idShapefile, (IDispatch**)&_PolygonLayer);
		_PolygonLayer->CreateNew(CComBSTR(""), ShpfileType::SHP_POLYGON, &vb);
		if (vb == VARIANT_TRUE)
		{
			_PolygonLayer->put_Volatile(VARIANT_TRUE);
			_PolygonLayer->put_Selectable(VARIANT_TRUE);
			// add layer to map
			layerHandle = _PolygonLayerHandle = this->AddLayer((LPDISPATCH)_PolygonLayer, Visible);
			// rendering?
			this->SetShapeLayerLineColor(_PolygonLayerHandle, 0xFFFF00);
			this->SetShapeLayerFillColor(_PolygonLayerHandle, 0xFFFF00);
			this->SetShapeLayerFillTransparency(_PolygonLayerHandle, .30);
		}
		break;
	default:
		break;
	}
	if (vb == VARIANT_TRUE)
	{
		// one-time set up of projection for transformations
		if (!_WGS84)
		{
			ComHelper::CreateInstance(idGeoProjection, (IDispatch**)&_WGS84);
			_WGS84->SetWgs84(&vb);
			_WGS84->StartTransform(this->GetGeoProjection(), &vb);
		}

		return layerHandle;
	}
	// if we get here, something went wrong
	return -1;
}

// ***************************************************************
//		AddVolatilePoint()
// ***************************************************************
BSTR CMapView::AddVolatilePoint(DOUBLE Lon, DOUBLE Lat)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (!_PointLayer)
	{
		CComBSTR msg("Volatile Point Layer does not yet exist");
		return msg;
	}

	VARIANT_BOOL vb;
	long idx;
	// create Point Shape
	CComPtr<IShape> pShape = NULL;
	ComHelper::CreateShape(&pShape);
	pShape->Create(ShpfileType::SHP_POINT, &vb);
	// set lon, lat for creation of WKT
	pShape->AddPoint(Lon, Lat, &idx);
	CComBSTR bstrWKT;
	pShape->ExportToWKT(&bstrWKT);
	// now reproject to map projection
	_WGS84->Transform(&Lon, &Lat, &vb);
	// update point shape
	pShape->put_XY(0, Lon, Lat, &vb);
	// add to Layer
	_PointLayer->EditAddShape(pShape, &idx);
	_PointLayer->StopEditingShapes(VARIANT_TRUE, VARIANT_TRUE, NULL, &vb);
	// return string with point_handle, WKT of point
	CAtlString strResult;
	strResult.Format("%d%s%s", idx, "\037", CAtlString(bstrWKT));
	CComBSTR bstrResult((LPCTSTR)strResult);
	return bstrResult;
}

// ***************************************************************
//		RemoveVolatileGeometry()
// ***************************************************************
void CMapView::RemoveVolatileGeometry(LONG LayerHandle, LONG GeomHandle)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	VARIANT_BOOL vb;
	LONG count;

	IShapefile* sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// layer must exist and point index must exist
		if (sf && (sf->get_NumShapes(&count) == S_OK) && (count > GeomHandle))
		{
			sf->EditDeleteShape(GeomHandle, &vb);
			if (vb == VARIANT_TRUE)
				sf->StopEditingShapes(VARIANT_TRUE, VARIANT_TRUE, NULL, &vb);
		}
	}
}

// ***************************************************************
//		AddVolatilePoint()
// ***************************************************************
BSTR CMapView::AddVolatileCircle(DOUBLE Lon, DOUBLE Lat, DOUBLE Radius)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (!_PolygonLayer)
	{
		CComBSTR msg("Volatile Polygon Layer does not yet exist");
		return msg;
	}

	VARIANT_BOOL vb;
	long idx;
	// create Polygon Shape
	CComPtr<IShape> pShape = NULL;
	ComHelper::CreateShape(&pShape);
	pShape->Create(ShpfileType::SHP_POLYGON, &vb);
	// radians conversion
	double radians = (2.0 * 3.14159265359 / 90.0);
	// circle will be a clockwise polygon of points
	for (int i = 0; i <= 90; i++)
	{
		pShape->AddPoint(Lon + Radius * cos(i * radians), Lat - Radius * sin(i * radians), &idx);
	}
	// reset the last point to the first point, since mathematically, they are not quite equal
	double x, y;
	pShape->get_XY(0, &x, &y, &vb);
	pShape->put_XY(90, x, y, &vb);

	CComBSTR bstrWKT;
	//pShape->ExportToWKT(&bstrWKT);
	//// now reproject to map projection
	//for (int i = 0; i <= 90; i++)
	//{
	//	pShape->get_XY(i, &x, &y, &vb);
	//	_WGS84->Transform(&x, &y, &vb);
	//	pShape->put_XY(i, x, y, &vb);
	//}

	// add to Layer
	_PolygonLayer->EditAddShape(pShape, &idx);
	_PolygonLayer->StopEditingShapes(VARIANT_TRUE, VARIANT_TRUE, NULL, &vb);
	// return string with point_handle, WKT of point
	CAtlString strResult;
	strResult.Format("%d%s%s", idx, "\037", CAtlString(bstrWKT));
	CComBSTR bstrResult((LPCTSTR)strResult);
	return bstrResult;
}

// ***************************************************************
//		GetLayerFeatureByGeometry()
// ***************************************************************
BSTR CMapView::GetLayerFeatureByGeometry(LONG SearchLayerHandle, LONG VolatileLayerHandle, LONG VolatileGeomHandle)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CAtlString strResult;
	tkSpatialRelation relation = tkSpatialRelation::srIntersects;

	// make sure we've got the Search layer
	CComPtr<IShapefile> sLayer = this->GetShapefile(SearchLayerHandle);
	if (!sLayer)
	{
		strResult.Format("Search layer with handle = '%d' not found", SearchLayerHandle);
	}
	else
	{
		// make sure we've got the Volatile geometry
		CComPtr<IShapefile> vLayer = this->GetShapefile(VolatileLayerHandle);
		if (!vLayer)
		{
			strResult.Format("Volatile layer with handle = '%d' not found", VolatileLayerHandle);
		}
		else
		{
			// and make sure we've got the search Geometry
			CComPtr<IShape> vShape = NULL;
			vLayer->get_Shape(VolatileGeomHandle, &vShape);
			if (!vShape)
			{
				strResult.Format("Geometry with handle = '%d' not found", VolatileGeomHandle);
			}
			else
			{
				long idx;
				VARIANT_BOOL vb;
				ShpfileType sfType;
				CComPtr<IShape> bufferedShape = NULL;
				CComPtr<IShapefile> defLayer;
				// create 'definition' layer
				ComHelper::CreateInstance(idShapefile, (IDispatch**)&defLayer);
				defLayer->CreateNew(L"", ShpfileType::SHP_POLYGON, &vb);
				defLayer->StartEditingShapes(VARIANT_TRUE, NULL, &vb);
				// if definition geometry is a point, create a buffer around it
				if ((vShape->get_ShapeType(&sfType) == S_OK) && (sfType == ShpfileType::SHP_POINT))
				{
					// add buffer
					vShape->Buffer(SearchTolerance, 180, &bufferedShape);
					// add to definition layer
					defLayer->EditAddShape(bufferedShape, &idx);
				}
				else
				{
					// else just add 2D geometry to layer
					defLayer->EditAddShape(vShape, &idx);
				}
				// save to definition layer
				defLayer->StopEditingShapes(VARIANT_TRUE, VARIANT_TRUE, NULL, &vb);

				CComVariant sResults;
				//// relation depends on searched geometry
				//ShpfileType sShapeType;
				//sLayer->get_ShapefileType2D(&sShapeType);
				//switch (sShapeType)
				//{
				//case ShpfileType::SHP_POINT:
				//	break;
				//case ShpfileType::SHP_POLYLINE:
				//	break;
				//case ShpfileType::SHP_POLYGON:
				//	break;
				//default:
				//	break;
				//}
				// try the select
				sLayer->SelectByShapefile(defLayer, tkSpatialRelation::srIntersects, VARIANT_FALSE, &sResults, NULL, &vb);
				if (vb == VARIANT_TRUE)
				{
					long* pData;
					long uBound, lBound;
					SafeArrayAccessData(sResults.parray, (void**)&pData);
					SafeArrayGetLBound(sResults.parray, 1, &lBound);
					SafeArrayGetUBound(sResults.parray, 1, &uBound);
					for (int i = lBound; i <= uBound; i++)
					{
						// iterate list of field indexes
						CComVariant sValue;
						for (long idx = 0; idx < (_featureColumns[SearchLayerHandle]).size(); idx++)
						{
							sLayer->get_CellValue((_featureColumns[SearchLayerHandle]).at(idx), (long)(pData[i]), &sValue);
							CAtlString val(sValue);
							if (strResult.GetLength() == 0)
							{
								strResult = val;
							}
							else
							{
								strResult.Format("%s%s%s", strResult, "\037", val);
							}
						}
						// if not on last entry, add ch30 separator
						if (i < uBound) strResult.Append("\036");
					}
				}
			}
		}
	}
	// return results
	//CComBSTR bstrResult((LPCTSTR)strResult);
	//return bstrResult;
	return strResult.AllocSysString();
	//return A2BSTR((LPCTSTR)strResult);
}

