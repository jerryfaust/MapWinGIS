// Swift-specific methods added for custom behavior

#pragma once
#include "stdafx.h"
#include "Map.h"
#include "ComHelpers\ProjectionHelper.h"

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

	// save feature columns
	_featureColumns[LayerHandle] = CString(ColumnName);
}


// ****************************************************************** 
//		SetLayerLabelColumn
// ****************************************************************** 
void CMapView::SetLayerLabelColumn(LONG LayerHandle, LPCTSTR ColumnName)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// save label columns
	_labelColumns[LayerHandle] = CString(ColumnName);
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
	CString strColumn = _labelColumns[LayerHandle];
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
CString BuildLabelExpression(CString input)
{
	CString exp;
	int tokenPos = 0;
	CString strToken = input.Tokenize("\037", tokenPos);
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
			CString strPlus = strToken.Tokenize("+", plusPos);
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
			CString strColumn = _labelColumns[LayerHandle];
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
				CString strExpr = BuildLabelExpression(strColumn);
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
					CString sfName(Filename);
					CString newName(sfName);
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
	CString strDirectory(ConnectionString);
	int pos = strDirectory.ReverseFind('\\');
	strDirectory = strDirectory.Left(pos);
	CComBSTR bstrShapefile((LPCTSTR)strDirectory);

	// if placing in GDB directory
	//CComBSTR bstrShapefile(ConnectionString);

	bstrShapefile.Append("\\");
	bstrShapefile.Append(bstrTable);
	bstrShapefile.Append(".shp");

	// does the Shapefile already exist?
	handle = AddLayerFromFilename(CString(bstrShapefile), tkFileOpenStrategy::fosAutoDetect, visible);
	if (handle >= 0) return handle;

	// else continue on loading from database
	_fileManager->OpenFromDatabase(bstrConnection, bstrTable, &layer);
	if (layer)
	{
		VARIANT_BOOL vbResult;
		layer->get_SupportsEditing(tkOgrSaveType::ostSaveAll, &vbResult);
		// for now, only if a GDB datasource
		CString strConnStr(ConnectionString);
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

	CComBSTR result;

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

	return result;
}

// volatile layer references
CComPtr<IShapefile> _PointLayer = NULL;
long _PointLayerHandle = -1;
CComPtr<IGeoProjection> _WGS84 = NULL;

// ***************************************************************
//		SetupVolatileLayers()
// ***************************************************************
VARIANT_BOOL CMapView::SetupVolatileLayers()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	VARIANT_BOOL vb;
	// set up Point layer
	ComHelper::CreateInstance(idShapefile, (IDispatch**)&_PointLayer);
	_PointLayer->CreateNew(CComBSTR(""), ShpfileType::SHP_POINT, &vb);
	if (vb == VARIANT_TRUE)
	{
		_PointLayer->put_Volatile(VARIANT_TRUE);
		_PointLayer->put_Selectable(VARIANT_TRUE);
		// set up projection for transformations
		//CComPtr<IGeoProjection> gpWGS84 = NULL;
		ComHelper::CreateInstance(idGeoProjection, (IDispatch**)&_WGS84);
		_WGS84->SetWgs84(&vb);
		_WGS84->StartTransform(this->GetGeoProjection(), &vb);
		//_PointLayer->put_GeoProjection(gp);
		// add to map
		_PointLayerHandle = this->AddLayer((LPDISPATCH)_PointLayer, TRUE);
		// rendering?
		this->SetShapeLayerPointColor(_PointLayerHandle, 255);
		this->SetShapeLayerPointSize(_PointLayerHandle, 10);

		return VARIANT_TRUE;
	}
	// if we get here, something went wrong
	return VARIANT_FALSE;
}

// ***************************************************************
//		AddVolatilePoint()
// ***************************************************************
LONG CMapView::AddVolatilePoint(DOUBLE lon, DOUBLE lat)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// create layer if necessary
	if (!_PointLayer && SetupVolatileLayers() == VARIANT_FALSE)
	{
		// could not create the layers
		return -1;
	}

	VARIANT_BOOL vb;
	long idx;
	// reproject
	_WGS84->Transform(&lon, &lat, &vb);
	// create Point Shape
	CComPtr<IShape> pShape = NULL;
	ComHelper::CreateShape(&pShape);
	pShape->Create(ShpfileType::SHP_POINT, &vb);
	pShape->AddPoint(lon, lat, &idx);
	CComBSTR bstrWKT;
	pShape->ExportToWKT(&bstrWKT);
	// add to Layer
	_PointLayer->EditAddShape(pShape, &idx);
	_PointLayer->StopEditingShapes(VARIANT_TRUE, VARIANT_TRUE, NULL, &vb);
	// return handle (points are 100-based)
	return (100 + idx);
}
