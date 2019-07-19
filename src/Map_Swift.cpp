// Swift-specific methods added for custom behavior

#pragma once
#include "stdafx.h"
#include "Map.h"
#include "ComHelpers\ProjectionHelper.h"
#include "GeosConverter.h"
#include "GeosHelper.h"

enum volatileLayerType
{
	vltPolygon = 1,
	vltPolyline,
	vltPoint
};

// split characters
#define ch29 "\035"
#define ch30 "\036"
#define ch31 "\037"


// volatile layer references
long _PointLayerHandle = -1;
long _PolylineLayerHandle = -1;
long _PolygonLayerHandle = -1;
CComPtr<IShapefile> _PointLayer = NULL;
CComPtr<IShapefile> _PolylineLayer = NULL;
CComPtr<IShapefile> _PolygonLayer = NULL;

// reprojector, from WGS84 to the current map projection
CComPtr<IGeoProjection> _WGS84 = NULL;
// reprojector, from map projection to WGS84
CComPtr<IGeoProjection> _MapProj = NULL;

double SearchTolerance = 10.0;
// minimum point diamter for zoom
long _PointDiameter = 100;
// current measuring type (may be custom value)
long _MeasuringType = 0;
long _MeasureLineColor = 0;

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


std::vector<CAtlString> CMapView::ParseDelimitedStrings(LPCTSTR strings, LPCTSTR delimiter)
{
	std::vector<CAtlString> results;
	// get delimited string into CString
	CAtlString strStrings(strings);
	// we have to have something...
	if (strStrings.GetLength() > 0)
	{
		// is it a single or multi-field string?
		if (strStrings.Find(delimiter) < 0)
		{
			results.push_back(strStrings);
		}
		else
		{
			// multi-field string
			int tokenPos = 0;
			CAtlString strToken = strStrings.Tokenize(delimiter, tokenPos);
			while (!strToken.IsEmpty())
			{
				results.push_back(strToken);
				// next?
				strToken = strStrings.Tokenize(delimiter, tokenPos);
			}
		}
	}
	return results;
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
		_featureColumns[LayerHandle].clear();
		// parse column names
		std::vector<CAtlString> columns = ParseDelimitedStrings(ColumnName, ch31);
		// get field indices by their names
		for (CAtlString col : columns)
		{
			CComBSTR bstrColName(col);
			sf->get_FieldIndexByName(bstrColName.Copy(), &fieldIndex);
			_featureColumns[LayerHandle].push_back(fieldIndex);
		}
		//// save feature columns (unpack into field indices
		//CAtlString strColumns(ColumnName);
		//// is it a single or multi-field label?
		//if (strColumns.Find(ch31) < 0)
		//{
		//	// single-field labeling
		//	CComBSTR bstrColName(strColumns);
		//	sf->get_FieldIndexByName(bstrColName.Copy(), &fieldIndex);
		//	_featureColumns[LayerHandle].push_back(fieldIndex);
		//}
		//else
		//{
		//	// multi-field labeling
		//	int tokenPos = 0;
		//	CAtlString strToken = strColumns.Tokenize(ch31, tokenPos);
		//	while (!strToken.IsEmpty())
		//	{
		//		CComBSTR bstrColName(strToken);
		//		sf->get_FieldIndexByName(bstrColName.Copy(), &fieldIndex);
		//		_featureColumns[LayerHandle].push_back(fieldIndex);
		//		// next?
		//		strToken = strColumns.Tokenize(ch31, tokenPos);
		//	}
		//}
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
			//CComBSTR bstrColumn((LPCTSTR)strColumn);
			//sf->put_SortField(bstrColumn);
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
					//labels->put_FrameVisible(VARIANT_TRUE);
					//labels->put_FrameType(tkLabelFrameType::lfRectangle);
					//labels->put_FrameOutlineColor(0);
					//labels->put_InboxAlignment(tkLabelAlignment::laBottomCenter);
					// labels
					labels->put_Alignment(tkLabelAlignment::laBottomCenter);
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
					labels->put_Alignment(tkLabelAlignment::laCenter);
					labels->put_AutoOffset(VARIANT_FALSE);
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
	CComPtr<IShapefile> sf = GetShapefile(LayerHandle);
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

			//labels->put_LogScaleForSize(VARIANT_TRUE);
			//labels->put_UseVariableSize(VARIANT_TRUE);
			//labels->put_FontSize2(FontSize + 6);

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
	CComPtr<IShapefile> sf = GetShapefile(LayerHandle);
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
//		SetLayerLabelFrame
// ****************************************************************** 
void CMapView::SetLayerLabelFrame(LONG LayerHandle, LONG BackColor, LONG FrameColor)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// get Shapefile reference
	CComPtr<IShapefile> sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// Labels reference
		ILabels* labels = nullptr;
		sf->get_Labels(&labels);
		if (labels != nullptr)
		{
			labels->put_FrameVisible(VARIANT_TRUE);
			labels->put_FrameBackColor(BackColor);
			labels->put_FrameOutlineColor(FrameColor);
			labels->put_InboxAlignment(tkLabelAlignment::laBottomCenter);

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
	CComPtr<IShapefile> sf = GetShapefile(LayerHandle);
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
	CAtlString strToken = input.Tokenize(ch31, tokenPos);
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
		strToken = input.Tokenize(ch31, tokenPos);
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
	CComPtr<IShapefile> sf = GetShapefile(LayerHandle);
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
			if (strColumn.Find(ch31) < 0)
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

void DeleteShapefile(LPCTSTR ShapefileName)
{
	CAtlString sfName(ShapefileName);
	sfName.MakeLower();
	remove((LPCTSTR)sfName);
	sfName.Replace(".shp", ".shx");
	remove((LPCTSTR)sfName);
	sfName.Replace(".shx", ".dbf");
	remove((LPCTSTR)sfName);
	sfName.Replace(".dbf", ".prj");
	remove((LPCTSTR)sfName);
	Sleep(100);
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
					DeleteShapefile(Filename);

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
long CMapView::AddDatasourceAndResave(LPCTSTR ConnectionString, LPCTSTR TableName, LPCTSTR Columns, BOOL visible)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	USES_CONVERSION;
	long handle = -1;
	IOgrLayer* layer = NULL;
	CComBSTR bstrConnection(ConnectionString);
	CComBSTR bstrTable(TableName);
	CComBSTR bstrColumns(Columns);

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

	// if no columns are specified, and a Shapefile already exists, just load it and exit
	if (bstrColumns.Length() == 0)
	{
		// does the Shapefile already exist?
		handle = AddLayerFromFilename(CAtlString(bstrShapefile), tkFileOpenStrategy::fosAutoDetect, visible);
		if (handle >= 0) return handle;
	}

	// else continue on loading from database
	layer = OpenLayerWithColumns(bstrConnection, bstrTable, bstrColumns);
	//_fileManager->OpenFromDatabase(bstrConnection, bstrTable, &layer);
	if (layer)
	{
		//VARIANT_BOOL vbResult;
		//layer->get_SupportsEditing(tkOgrSaveType::ostSaveAll, &vbResult);
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
					// pre-delete any previously saved files
					DeleteShapefile((LPCTSTR)CString(bstrShapefile));
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

// Add layer with only those columns specified, or entire layer if no columns specified
IOgrLayer* CMapView::OpenLayerWithColumns(BSTR strConnection, BSTR strTable, BSTR strColumns)
{
	long handle = -1;
	IOgrLayer* layer = NULL;
	CComBSTR bstrTable(strTable);
	CAtlString cstrColumns(strColumns);
	CAtlString sqlQuery;

	// do we have columns specified ?
	if (cstrColumns.GetLength() > 0)
	{
		// build SQL query
		sqlQuery = "SELECT ";
		bool firstCol = true;
		vector<CAtlString> columns = ParseDelimitedStrings((LPCTSTR)cstrColumns, ch31);
		for (CAtlString col : columns)
		{
			if (firstCol)
			{
				sqlQuery.Append(col);
				firstCol = false;
			}
			else
			{
				sqlQuery.AppendFormat(", %s", col);
			}
		}
		// FROM
		sqlQuery.AppendFormat(" FROM %s", (LPCTSTR)CString(strTable));
		// override table with query
		bstrTable = (LPCTSTR)sqlQuery;
	}

	// open the table/query
	_fileManager->OpenFromDatabase(strConnection, bstrTable, &layer);
	// return layer reference
	return layer;
}

// ***************************************************************
//		QueryLayer()
// ***************************************************************
BSTR CMapView::QueryLayer(LONG LayerHandle, LPCTSTR WhereClause)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CAtlString strResult("");

	// make sure we've got the Search layer
	CComPtr<IShapefile> sLayer = GetShapefile(LayerHandle);
	if (!sLayer)
	{
		strResult.Format("Query layer with handle = '%d' not found", LayerHandle);
	}
	else
	{
		ITable* tab = nullptr;
		sLayer->get_Table(&tab);
		if (tab)
		{
			CComBSTR expression(WhereClause);
			CComBSTR errorString;
			CComVariant results;
			VARIANT_BOOL vb;
			tab->Query(expression, &results, &errorString, &vb);
			if (vb != VARIANT_FALSE)
			{
				// get user-specified column values
				long* pData;
				long uBound, lBound;
				SafeArrayAccessData(results.parray, (void**)&pData);
				SafeArrayGetLBound(results.parray, 1, &lBound);
				SafeArrayGetUBound(results.parray, 1, &uBound);
				for (int i = lBound; i <= uBound; i++)
				{
					// current shape index (GeomHandle)
					long shapeIdx = (long)(pData[i]);
					// iterate list of field indexes
					CComVariant sValue;
					long numFields = (_featureColumns[LayerHandle]).size();
					if (numFields == 0)
					{
						CAtlString gh;
						gh.Format("%d", shapeIdx);
						// no fields, just set GeomHandle
						if (strResult.GetLength() == 0)
						{
							// first entry
							strResult = gh;
						}
						else
						{
							// append start of feature
							strResult.Append(gh);
						}
					}
					else
					{
						for (long idx = 0; idx < numFields; idx++)
						{
							long fieldIdx = (_featureColumns[LayerHandle]).at(idx);
							// get next value, convert to string
							sLayer->get_CellValue(fieldIdx, shapeIdx, &sValue);
							CAtlString val(sValue);
							// on first row entry of feature, prepend GeomHandle
							if (idx == 0)
							{
								// prepend GeomHandle to first field value
								if (strResult.GetLength() == 0)
								{
									// very first entry
									strResult.Format("%d%s%s", shapeIdx, ch31, val);
								}
								else
								{
									// append start of feature
									strResult.Format("%s%d%s%s", strResult, shapeIdx, ch31, val);
								}
							}
							else
							{
								// intermediate field value
								strResult.Format("%s%s%s", strResult, ch31, val);
							}
						}
					}
					// if not on last entry, add ch30 separator
					if (i < uBound) strResult.Append(ch30);
				}
				// release SAFEARRAY lock
				SafeArrayUnaccessData(results.parray);
			}
		}
	}

	//return result;
	return strResult.AllocSysString();
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

// standard work to set up visibility for Deleted shapes
void CMapView::SetupLayerAttributes(IShapefile* sf, LPCTSTR Columns)
{
	//CComPtr<IShapeDrawingOptions> options;
	// projection
	CComPtr<IGeoProjection> gp = NULL;
	this->GetGeoProjection()->Clone(&gp);
	sf->put_GeoProjection(gp);
	// columns ?
	std::vector<CAtlString> cols = ParseDelimitedStrings(Columns, ch31);
	for (CAtlString col : cols)
	{
		// all will be String columns of 128 characters
		long idx;
		CComBSTR bCol(col);
		sf->EditAddField(bCol, FieldType::STRING_FIELD, 0, 128, &idx);
	}
	// attributes
	sf->put_Volatile(VARIANT_TRUE);
	sf->put_Selectable(VARIANT_TRUE);
	//// default visible
	//sf->get_DefaultDrawingOptions(&options);
	//options->put_Visible(VARIANT_TRUE);
}

// ***************************************************************
//		AddUserLayer()
// ***************************************************************
LONG CMapView::AddUserLayer(LONG GeometryType, LPCTSTR Columns, BOOL Visible)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	VARIANT_BOOL vb;
	LONG layerHandle = -1;
	CComPtr<IShapefile> sf = NULL;

	// set up layer
	switch ((volatileLayerType)GeometryType)
	{
	case vltPoint:
		ComHelper::CreateInstance(idShapefile, (IDispatch**)&sf);
		sf->CreateNew(CComBSTR(""), ShpfileType::SHP_POINT, &vb);
		if (vb == VARIANT_TRUE)
		{
			SetupLayerAttributes(sf, Columns);
			// add layer to map
			layerHandle = this->AddLayer((LPDISPATCH)sf, Visible);
			// default rendering?
			this->SetShapeLayerPointColor(layerHandle, 255);
			this->SetShapeLayerPointSize(layerHandle, 14);
			this->SetShapeLayerPointType(layerHandle, 2); // star
		}
		break;
	case vltPolyline:
		ComHelper::CreateInstance(idShapefile, (IDispatch**)&sf);
		sf->CreateNew(CComBSTR(""), ShpfileType::SHP_POLYLINE, &vb);
		if (vb == VARIANT_TRUE)
		{
			SetupLayerAttributes(sf, Columns);
			// add layer to map
			layerHandle = this->AddLayer((LPDISPATCH)sf, Visible);
		}
		break;
	case vltPolygon:
		ComHelper::CreateInstance(idShapefile, (IDispatch**)&sf);
		sf->CreateNew(CComBSTR(""), ShpfileType::SHP_POLYGON, &vb);
		if (vb == VARIANT_TRUE)
		{
			//FillColor = RGB(255, 165, 0);
			//FillTransparency = 100;
			//LineColor = RGB(255, 127, 0);
			//LineWidth = 2.0f;
			SetupLayerAttributes(sf, Columns);
			// add layer to map
			layerHandle = this->AddLayer((LPDISPATCH)sf, Visible);
			// default rendering?
			this->SetShapeLayerLineWidth(layerHandle, 2);
			// colors to match measuring tool
			this->SetShapeLayerLineColor(layerHandle, _MeasureLineColor); // 0xFFFF00); // RGB(255, 127, 0)
			this->SetShapeLayerFillColor(layerHandle, _MeasureLineColor); // 0xFFFF00); // RGB(255, 165, 0)
			this->SetShapeLayerFillTransparency(layerHandle, 0.30f);
		}
		break;
	default:
		break;
	}
	if (vb == VARIANT_TRUE)
	{
		// one-time set up of projections for transformations
		if (!_WGS84)
		{
			ComHelper::CreateInstance(idGeoProjection, (IDispatch**)&_WGS84);
			_WGS84->SetWgs84(&vb);
			_WGS84->StartTransform(this->GetGeoProjection(), &vb);
		}
		if (!_MapProj)
		{
			ComHelper::CreateInstance(idGeoProjection, (IDispatch**)&_MapProj);
			this->GetGeoProjection()->Clone(&_MapProj);
			_MapProj->StartTransform(_WGS84, &vb);

			//this->GetMeasuring()->put_MeasuringType(tkMeasuringType::MeasureDistance);
			this->GetMeasuring()->put_LengthUnits(tkLengthDisplayMode::ldmAmerican);
			this->GetMeasuring()->put_AreaUnits(tkAreaDisplayMode::admAmerican);
			this->GetMeasuring()->put_Persistent(VARIANT_TRUE);
			//this->GetMeasuring()->put_MeasuringType(tkMeasuringType::MeasureArea);
		}

		return layerHandle;
	}
	// if we get here, something went wrong
	return -1;
}

// ***************************************************************
//		AddUserPoint()
// ***************************************************************
BSTR CMapView::AddUserPoint(DOUBLE xCoord, DOUBLE yCoord)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (!_PointLayer)
	{
		CComBSTR msg("Volatile Point Layer does not yet exist");
		return msg;
	}

	VARIANT_BOOL vb;
	long idx = -1;
	// create Point Shape
	CComPtr<IShape> pShape = NULL;
	ComHelper::CreateShape(&pShape);
	pShape->Create(ShpfileType::SHP_POINT, &vb);
	// set specified lon, lat 
	pShape->AddPoint(xCoord, yCoord, &idx);
	//CComBSTR bstrWKT;
	//pShape->ExportToWKT(&bstrWKT);
	// now reproject to map projection
	_WGS84->Transform(&xCoord, &yCoord, &vb);
	// update point shape
	pShape->put_XY(0, xCoord, yCoord, &vb);
	// add to Layer
	idx = -1;
	_PointLayer->StartEditingShapes(VARIANT_TRUE, NULL, &vb);
	_PointLayer->EditAddShape(pShape, &idx);
	_PointLayer->StopEditingShapes(VARIANT_TRUE, VARIANT_TRUE, NULL, &vb);
	// return string with point_handle, WKT of point
	CAtlString strResult;
	//strResult.Format("%d%s%s", idx, ch31, CAtlString(bstrWKT));
	strResult.Format("%d", idx);
	//
	return strResult.AllocSysString();
}

// ***************************************************************
//		AddUserCircle()
// ***************************************************************
BSTR CMapView::AddUserCircle(DOUBLE xCoord, DOUBLE yCoord, DOUBLE Radius)
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
		pShape->AddPoint(xCoord + Radius * cos(i * radians), yCoord - Radius * sin(i * radians), &idx);
	}
	// reset the last point to the first point, since mathematically, they are not quite equal
	double x, y;
	pShape->get_XY(0, &x, &y, &vb);
	pShape->put_XY(90, x, y, &vb);

	// add to Layer
	idx = -1;
	_PolygonLayer->StartEditingShapes(VARIANT_TRUE, NULL, &vb);
	_PolygonLayer->EditAddShape(pShape, &idx);
	_PolygonLayer->StopEditingShapes(VARIANT_TRUE, VARIANT_TRUE, NULL, &vb);
	//// return string with point_handle, WKT of point
	//CComBSTR bstrWKT;
	//// now reproject to WGS84 (to new shape)
	//CComPtr<IShape> wgsShape = NULL;
	//ComHelper::CreateShape(&wgsShape);
	//wgsShape->Create(ShpfileType::SHP_POLYGON, &vb);
	//long newIdx = -1;
	//for (int i = 0; i <= 90; i++)
	//{
	//	// take from original shape
	//	pShape->get_XY(i, &x, &y, &vb);
	//	_MapProj->Transform(&x, &y, &vb);
	//	// add to new shape (throw-away index)
	//	wgsShape->AddPoint(x, y, &newIdx);
	//}
	//wgsShape->ExportToWKT(&bstrWKT);

	CAtlString strResult;
	//strResult.Format("%d%s%s", idx, ch31, CAtlString(bstrWKT));
	strResult.Format("%d", idx);
	return strResult.AllocSysString();
}

// ***************************************************************
//		RemoveUserGeometry()
// ***************************************************************
void CMapView::RemoveUserGeometry(LONG LayerHandle, LONG GeomHandle)
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
				// if definition geometry is a point or a line, create a buffer around it
				if ((vShape->get_ShapeType2D(&sfType) == S_OK) && 
					(sfType == ShpfileType::SHP_POINT || sfType == ShpfileType::SHP_POLYLINE))
				{
					// add buffer
					vShape->Buffer(SearchTolerance, 16, &bufferedShape);
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
						// current shape index (GeomHandle)
						long shapeIdx = (long)(pData[i]);
						// iterate list of field indexes
						CComVariant sValue;
						long numFields = (_featureColumns[SearchLayerHandle]).size();
						if (numFields == 0)
						{
							CAtlString gh;
							gh.Format("%d", shapeIdx);
							// no fields, just set GeomHandle
							if (strResult.GetLength() == 0)
							{
								// first entry
								strResult = gh;
							}
							else
							{
								// append start of feature
								strResult.Append(gh);
							}
						}
						else
						{
							for (long idx = 0; idx < numFields; idx++)
							{
								long fieldIdx = (_featureColumns[SearchLayerHandle]).at(idx);
								// get next value, convert to string
								sLayer->get_CellValue(fieldIdx, shapeIdx, &sValue);
								CAtlString val(sValue);
								// on first row entry of feature, prepend GeomHandle
								if (idx == 0)
								{
									// prepend GeomHandle to first field value
									if (strResult.GetLength() == 0)
									{
										// very first entry
										strResult.Format("%d%s%s", shapeIdx, ch31, val);
									}
									else
									{
										// append start of feature
										strResult.Format("%s%d%s%s", strResult, shapeIdx, ch31, val);
									}
								}
								else if (idx < numFields)
								{
									// intermediate field value
									strResult.Format("%s%s%s", strResult, ch31, val);
								}
								else
								{
									// last field value
									strResult.Append(val);
								}
							}
						}
						// if not on last entry, add ch30 separator
						if (i < uBound) strResult.Append(ch30);
					}
					// release SAFEARRAY lock
					SafeArrayUnaccessData(sResults.parray);
				}
			}
		}
	}

	// PROGRESS IDE cannot accept more than 32K string
	if (strResult.GetLength() > 32000)
	{
		// lop off at the last ch30 prior to 32K
		strResult = strResult.Left(32000);
		strResult = strResult.Left(strResult.ReverseFind(ch30[0]));
		//strResult = "NOK";
	}

	// return results
	return strResult.AllocSysString();
}

// ***************************************************************
//		SetVolatileLayer()
// ***************************************************************
void CMapView::SetVolatileLayer(LONG GeometryType, LONG LayerHandle)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CComPtr<IShapefile> sf = NULL;

	// set up layer
	switch ((volatileLayerType)GeometryType)
	{
	case vltPoint:
		// save layer handle
		_PointLayerHandle = LayerHandle;
		// get layer reference
		_PointLayer = this->GetShapefile(LayerHandle);
		break;
	case vltPolyline:
		// save layer handle
		_PolylineLayerHandle = LayerHandle;
		// get layer reference
		_PolylineLayer = this->GetShapefile(LayerHandle);
		break;
	case vltPolygon:
		// save layer handle
		_PolygonLayerHandle = LayerHandle;
		// get layer reference
		_PolygonLayer = this->GetShapefile(LayerHandle);
		break;
	default:
		break;
	}
}

// ***************************************************************
//		PlaceGeometryByWKT()
// ***************************************************************
LONG CMapView::PlaceGeometryByWKT(LONG LayerHandle, LPCTSTR WKT)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	long idx = -1;
	// get layer reference
	CComPtr<IShapefile> sf = this->GetShapefile(LayerHandle);
	if (sf)
	{
		VARIANT_BOOL vb;
		// create a Shape
		CComPtr<IShape> shp = NULL;
		ComHelper::CreateShape(&shp);
		// fill the shape based on the WKT
		CComBSTR bstr(WKT);
		shp->ImportFromWKT(bstr, &vb);
		// ok?
		if (vb == VARIANT_TRUE)
		{
			// transform to the map projection
			long count;
			double x, y;
			shp->get_NumPoints(&count);
			for (int i = 0; i < count; i++)
			{
				shp->get_XY(i, &x, &y, &vb);
				_WGS84->Transform(&x, &y, &vb);
				shp->put_XY(i, x, y, &vb);
			}
			// add the shape to the layer
			if (sf->StartEditingShapes(VARIANT_TRUE, NULL, &vb) == S_OK && vb == VARIANT_TRUE)
			{
				sf->EditAddShape(shp, &idx);
				sf->StopEditingShapes((idx >= 0) ? VARIANT_TRUE : VARIANT_FALSE, (idx >= 0) ? VARIANT_TRUE : VARIANT_FALSE, NULL, &vb);
			}
		}
	}
	// return shape index
	return idx;
}

// ***************************************************************
//		CopyGeometryByHandle()
// ***************************************************************
LONG CMapView::CopyGeometryByHandle(LONG SourceLayerHandle, LONG SourceGeomHandle, LONG TargetLayerHandle)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	LONG TargetGeomHandle = -1;
	// get source layer reference
	CComPtr<IShapefile> source = this->GetShapefile(SourceLayerHandle);
	// get target layer reference
	CComPtr<IShapefile> target = this->GetShapefile(TargetLayerHandle);
	if (source && target)
	{
		VARIANT_BOOL vb;
		// get geometry from source
		CComPtr<IShape> shp = NULL;
		source->get_Shape(SourceGeomHandle, &shp);
		if (shp)
		{
			// insert into target layer
			if (target->StartEditingShapes(VARIANT_TRUE, NULL, &vb) == S_OK && vb == VARIANT_TRUE)
			{
				CComPtr<IShape> newShp = NULL;
				// test
				//shp->Buffer(20, 180, &newShp);
				// set copy into target layer
				shp->Clone(&newShp);
				target->EditAddShape(newShp, &TargetGeomHandle);
				target->StopEditingShapes((TargetGeomHandle < 0) ? VARIANT_FALSE : VARIANT_TRUE, (TargetGeomHandle < 0) ? VARIANT_FALSE : VARIANT_TRUE, NULL, &vb);
			}
		}
	}
	// return target index (as geom handle)
	return TargetGeomHandle;
}


// ****************************************************************** 
//		SetCellValues
// ****************************************************************** 
void CMapView::SetCellValues(LONG LayerHandle, LONG GeomHandle, LPCTSTR NameValuePairs)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// set label values based on the specified name/value pairs
	VARIANT_BOOL vb;
	// get layer reference
	CComPtr<IShapefile> sf = this->GetShapefile(LayerHandle);
	if (sf && sf->StartEditingShapes(VARIANT_TRUE, NULL, &vb) == S_OK && vb == VARIANT_TRUE)
	{
		VARIANT_BOOL vb;
		// first split apart name/value pairs (character 30)
		std::vector<CAtlString> pairs = ParseDelimitedStrings(NameValuePairs, ch30);
		for (CAtlString pair : pairs)
		{
			long idx = -1;
			// now split each pair into name/value (character 31)
			std::vector<CAtlString> nameValue = ParseDelimitedStrings(pair, ch31);
			// name is nameValue[0], value is nameValue[1]
			CComBSTR name(nameValue[0]);
			CComVariant value((LPCTSTR)nameValue[1]);
			sf->get_FieldIndexByName(name, &idx);
			if (idx >= 0)
			{
				sf->EditCellValue(idx, GeomHandle, value, &vb);
				if (vb == VARIANT_FALSE) break;
			}
		}
		// save
		VARIANT_BOOL vbResult;
		sf->StopEditingShapes(vb, vb, NULL, &vbResult);
	}
}


// ***************************************************************
//		SetPointDiameter()
// ***************************************************************
void CMapView::SetPointDiameter(LONG Meters)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// save value
	_PointDiameter = Meters;
}


// ****************************************************************** 
//		ZoomToGeometry
// ****************************************************************** 
double CMapView::ZoomToGeometry(LONG LayerHandle, LONG GeomHandle)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	double scale = this->GetCurrentScale();

	// get Shapefile reference
	CComPtr<IShapefile> sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// get Shape
		CComPtr<IShape> shp = NULL;
		sf->get_Shape(GeomHandle, &shp);
		if (shp)
		{
			// get Shape extents (special consideration for Points)
			CComPtr<IExtents> ext = NULL;
			//ShpfileType shapeType;
			//shp->get_ShapeType2D(&shapeType);
			//if (shapeType == ShpfileType::SHP_POINT)
			//{
			//	// create extents from point
			//	double x, y;
			//	VARIANT_BOOL vb;
			//	shp->get_XY(0, &x, &y, &vb);
			//	ComHelper::CreateExtents(&ext);
			//	ext->SetBounds(x - (_PointDiameter / 2.0), y - (_PointDiameter / 2.0), 0.0, x + (_PointDiameter / 2.0), y + (_PointDiameter / 2.0), 0.0);
			//}
			//else
			//{
				// get extents
				shp->get_Extents(&ext);
			//}
			// now zoom to extents
			//this->LockWindow(tkLockMode::lmLock);
			this->SetExtents(ext);
			//if (ZoomFactor > 1.0)
			//	this->ZoomOut(max(0.0, ZoomFactor - 1.0));
			// get new scale
			scale = this->GetCurrentScale();
			//// compare current scale with minimum setting
			//double minScale = this->GetLayerMinVisibleScale(LayerHandle);
			//if (scale < minScale)
			//{
			//	// bring it back out
			//	this->SetCurrentScale(minScale);
			//	scale = minScale;
			//}
			//this->LockWindow(tkLockMode::lmUnlock);
		}
	}
	// return current scale
	return scale;
}


// ***************************************************************
//		GetGeometryWKT()
// ***************************************************************
BSTR CMapView::GetGeometryWKT(LONG LayerHandle, LONG GeomHandle)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CComBSTR bstrWKT("");
	CString strResult;

	// get Shapefile reference
	CComPtr<IShapefile> sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// get Shape
		CComPtr<IShape> shp = NULL;
		sf->get_Shape(GeomHandle, &shp);
		if (shp)
		{
			VARIANT_BOOL vb;
			// create a copy for the transformed shape
			CComPtr<IShape> pShape = NULL;
			ComHelper::CreateShape(&pShape);
			// of same type as original type
			ShpfileType shpType;
			shp->get_ShapeType2D(&shpType);
			pShape->Create(shpType, &vb);
			// transform from map projection to WGS84
			long count, idx;
			double x, y;
			shp->get_NumPoints(&count);
			for (int i = 0; i < count; i++)
			{
				shp->get_XY(i, &x, &y, &vb);
				_MapProj->Transform(&x, &y, &vb);
				// add to shape copy
				pShape->AddPoint(x, y, &idx);
			}
			// get WKT string
			pShape->ExportToWKT(&bstrWKT);
		}
		else
		{
			strResult.Format("Geometry with handle = '%d' not found", GeomHandle);
			bstrWKT = (LPCTSTR)strResult;
		}
	}
	else
	{
		strResult.Format("Layer with handle = '%d' not found", LayerHandle);
		bstrWKT = (LPCTSTR)strResult;
	}

	// PROGRESS IDE cannot accept more than 32K string
#ifdef RELEASE_MODE
	if (bstrWKT.Length() > 32000)
		bstrWKT = "NOK";
#endif // RELEASE_MODE

	// return result
	//return bstrWKT;
	// large BSTR's don't always return properly,
	// but seem to work if set into CStrings and allocated out the door
	CAtlString cstr(bstrWKT);
	return cstr.AllocSysString();
}


// ***************************************************************
//		GetGeometryWKT()
// ***************************************************************
BSTR CMapView::GetGeometryWKTEx(LONG LayerHandle, LPCTSTR GeometryHandles)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CComBSTR bstrWKT("");
	CString strResult;

	// get Shapefile reference
	CComPtr<IShapefile> sf = GetShapefile(LayerHandle);
	if (sf)
	{
		VARIANT_BOOL vb;
		ShpfileType shpType;

		// create Shape to be union of all shapes
		CComPtr<IShape> shpUnion = NULL;
		//ComHelper::CreateShape(&shpUnion);
		//// of same type as original type
		sf->get_ShapefileType2D(&shpType);
		//// except points become a point collection
		//if (shpType == ShpfileType::SHP_POINT)
		//	shpType = ShpfileType::SHP_MULTIPOINT;
		//shpUnion->Create(shpType, &vb);

		// collection to hold all geometries
		vector<GEOSGeometry*> geoms;
		vector<IShape*> shapes;

		CComPtr<IShape> multiPartShape;
		ComHelper::CreateShape(&multiPartShape);
		multiPartShape->Create(shpType == ShpfileType::SHP_POINT ? ShpfileType::SHP_MULTIPOINT : shpType, &vb);

		// parse geometry handles
		std::vector<CAtlString> strHandles = ParseDelimitedStrings(GeometryHandles, ch30);
		// iterate
		for (CAtlString strHandle : strHandles)
		{
			// convert to numeric
			long handle = atol((LPCTSTR)strHandle);

			// get Shape
			CComPtr<IShape> shp = NULL;
			sf->get_Shape(handle, &shp);
			if (shp)
			{
				//// convert shape to GEOS geometry
				//GEOSGeometry* g = GeosConverter::ShapeToGeom(shp);
				//// add to collection
				//geoms.push_back(g);

				// current state of multi-part shape
				long partIdx, pointIdx;
				multiPartShape->get_NumParts(&partIdx);
				multiPartShape->get_NumPoints(&pointIdx);
				long numPoints;
				shp->get_NumPoints(&numPoints);
				for (int idx = 0; idx < numPoints; idx++)
				{
					double x, y;
					long newIdx;
					// get point from new shape
					shp->get_XY(idx, &x, &y, &vb);
					// add it to the multi-part line
					multiPartShape->AddPoint(x, y, &newIdx);
				}
				// identify the new part
				multiPartShape->InsertPart(pointIdx, &partIdx, &vb);
			}
		}

		long numParts;
		multiPartShape->get_NumParts(&numParts);

		// now, hopefully convert multi-part Shape to multi-part GEOS geometry
		GEOSGeometry* g = GeosConverter::ShapeToGeom(multiPartShape);
		// union the parts
		GEOSGeometry* geomUnion = GeosHelper::UnaryUnion(g);

		//// union all of the shapes
		//GEOSGeometry* geomUnion = GeosConverter::MergeGeometries(geoms, NULL, true, false);

		// convert back to a Shape
		vector<IShape*> vShapes;
		if (GeosConverter::GeomToShapes(geomUnion, &vShapes, false) && vShapes.size() > 0)
		{
			// hopefully there's only one
			int count = vShapes.size();
			ASSERT(count == 1);
			shpUnion = vShapes[0];
		}

		// transform from map projection to WGS84
		long count;
		double x, y;
		shpUnion->get_NumPoints(&count);
		for (int i = 0; i < count; i++)
		{
			shpUnion->get_XY(i, &x, &y, &vb);
			_MapProj->Transform(&x, &y, &vb);
			// replace the point
			shpUnion->put_XY(i, x, y, &vb);
		}
		// get WKT string
		shpUnion->ExportToWKT(&bstrWKT);
	}
	else
	{
		strResult.Format("Layer with handle = '%d' not found", LayerHandle);
		bstrWKT = (LPCTSTR)strResult;
	}

	// PROGRESS IDE cannot accept more than 32K string
#ifdef RELEASE_MODE
	if (bstrWKT.Length() > 32000)
		bstrWKT = "NOK";
#endif // RELEASE_MODE

	// return result
	//return bstrWKT;
	// large BSTR's don't always return properly,
	// but seem to work if set into CStrings and allocated out the door
	CAtlString cstr(bstrWKT);
	return cstr.AllocSysString();
}


// ***************************************************************
//		GetCurrentCenter()
// ***************************************************************
void CMapView::GetCurrentCenter(DOUBLE* Longitude, DOUBLE* Latitude)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// initialize for error
	*Longitude = 0.0;
	*Latitude = 0.0;

	// get current map extent
	CComPtr<IExtents> ext = this->GetExtents();
	// do we have something?
	if (ext)
	{
		double x, y;
		CComPtr<IPoint> center;
		// get the center point
		ext->get_Center(&center);
		if (center)
		{
			center->get_X(&x);
			center->get_Y(&y);
			// reproject to DMS
			VARIANT_BOOL vb;
			_MapProj->Transform(&x, &y, &vb);
			// success?
			if (vb == VARIANT_TRUE)
			{
				*Longitude = x;
				*Latitude = y;
			}
		}
	}
}


// *************************************************
//			LockWindowEx()						  
// *************************************************
void CMapView::LockWindowEx()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// increment the lock count
	_lockCount++;
}


// *************************************************
//			UnlockWindowEx()						  
// *************************************************
void CMapView::UnlockWindowEx(tkRedrawType RedrawType)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (_lockCount == 0)
	{
		return;    // no need to schedule more buffer reloads
	}

	// decrement the lock count
	_lockCount--;

	// are all locks released?
	if (_lockCount == 0)
	{
		RedrawCore(RedrawType, (RedrawType == tkRedrawType::RedrawSkipDataLayers) ? true : false);
	}
}


// *************************************************
//			SetMeasuringType()						  
// *************************************************
void CMapView::SetMeasuringType(LONG MeasuringType)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// set the setting in the measuring class
	// Type 2 is our custom circle measuring, which we handle with a restricted single-segment Distance measure
	this->GetMeasuring()->put_MeasuringType(MeasuringType == 2 ? tkMeasuringType::MeasureDistance : (tkMeasuringType)MeasuringType);
	this->SetCursorMode(tkCursorMode::cmMeasure);
	// measure type 2 restricts line drawing to one segment
	this->GetMeasuringBase()->SingleSegmentDistanceMeasure = (MeasuringType == 2);
	// save local value
	_MeasuringType = MeasuringType;
}

// *************************************************
//			ClearMeasuring()						  
// *************************************************
void CMapView::ClearMeasuring()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// clear all current measuring input
	this->GetMeasuring()->Clear();
	this->GetMeasuringBase()->SingleSegmentDistanceMeasure = false;
}

// *************************************************
//			RemoveLastMeasuringPoint()						  
// *************************************************
void CMapView::RemoveLastMeasuringPoint()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// undo the last-clicked measuring point
	CPoint point(_mouseX, _mouseY);
	// submit right-click with current mouse position
	this->OnRButtonDown(0, point);
}

// *************************************************
//			EndMeasuring()						  
// *************************************************
void CMapView::EndMeasuring(BOOL IncludeCurrentMousePosition)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// end measuring, submitting last point at current mouse position
	CPoint point(_mouseX, _mouseY);
	// is the SHIFT key pressed
	UINT kFlags = (GetKeyState(VK_SHIFT) & 0x8000) ? MK_SHIFT : 0;
	// for standard measure types (0 or 1) 
	// to include the mouse position, we have to submit a click
	if (_MeasuringType < 2 && IncludeCurrentMousePosition)
		this->OnLButtonDown(kFlags, point);
	// submit double-click with current mouse position
	this->OnLButtonDblClk(kFlags, point);
}


// *************************************************
//			SetMeasureLineColor()						  
// *************************************************
void CMapView::SetMeasureLineColor(LONG RgbColor)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// set the line color for the Measure display
	this->GetMeasuringBase()->LineColor = RgbColor;
	// save locally
	_MeasureLineColor = RgbColor;

	//CComPtr<IImage> pImg;
	//pImg = (IImage*)SnapShot(this->GetExtents());
	//VARIANT_BOOL vb;
	//CComBSTR bs("c:\\temp\\snapshot.jpg");
	//pImg->Save(bs, VARIANT_FALSE, ImageType::JPEG_FILE, NULL, &vb);
}


//// local helper function
//CComPtr<IShape> GetMeasureShape()
//{
//	VARIANT_BOOL vb;

//// local helper function
//CComPtr<IShape> GetMeasureShape()
//{
//	VARIANT_BOOL vb;
//
//	CComPtr<IShape> shp;
//	ComHelper::CreateShape(&shp);
//
//	int pointCount = this->GetMeasuringBase()->GetPointCount();
//	// bail out if no points
//	if (pointCount == 0)
//	{
//		this->GetMeasuring()->Clear();
//		return bstrWKT;
//	}
//
//	// special handling of custom Circle measuring (type = 2)
//	if (_MeasuringType == 2)
//	{
//		long idx;
//		// create Polygon Shape
//		shp->Create(ShpfileType::SHP_POLYGON, &vb);
//		// radians conversion
//		double radians = (2.0 * 3.14159265359 / 90.0);
//		// center point
//		double x, y;
//		MeasurePoint* mp = GetMeasuringBase()->GetPoint(0);
//		x = mp->x;
//		y = mp->y;
//		// radius
//		double radius = GetMeasureRadius();
//		// circle will be a clockwise polygon of points
//		for (int i = 0; i <= 90; i++)
//		{
//			shp->AddPoint(x + radius * cos(i * radians), y - radius * sin(i * radians), &idx);
//		}
//		// reset the last point to the first point, since mathematically, they are not quite equal
//		shp->get_XY(0, &x, &y, &vb);
//		shp->put_XY(90, x, y, &vb);
//	}
//	else
//	{
//		// else use measuring data
//		bool isPolygon;
//		// what do we have?
//		if (this->GetMeasuringBase()->HasPolygon(false))
//		{
//			shp->Create(ShpfileType::SHP_POLYGON, &vb);
//			isPolygon = true;
//		}
//		else if (this->GetMeasuringBase()->HasLine(false) && pointCount > 1)
//		{
//			shp->Create(ShpfileType::SHP_POLYLINE, &vb);
//			isPolygon = false;
//		}
//		else
//			// nothing, return empty WKT
//			return bstrWKT;
//
//		// iterate points in shape
//		for (int idx = 0; idx < pointCount; idx++)
//		{
//			long newIdx;
//			double x, y;
//			// transfer to shape points
//			MeasurePoint* pPoint = this->GetMeasuringBase()->GetPoint(idx);
//			shp->AddPoint(pPoint->x, pPoint->y, &newIdx);
//			// if on the first point...
//			if (idx == 0)
//			{
//				// save the first point
//				x = pPoint->x; y = pPoint->y;
//			}
//			// if on the last point of a polygon...
//			if (idx == (pointCount - 1) && isPolygon)
//			{
//				// we don't know if the polygon came from the line drawing or the area drawing,
//				// but if it came from the area, the last point is not equal to the first point
//				if (pPoint->x != x || pPoint->y != y)
//				{
//					// add new last point, being the same as first point
//					shp->AddPoint(x, y, &newIdx);
//				}
//			}
//		}
//	}
//
//	// get WKT
//	shp->ExportToWKT(&bstrWKT);
//
//	// return result
//	return bstrWKT;
//}


// ***************************************************************
//		GetMeasureWKT()
// ***************************************************************
BSTR CMapView::GetMeasureWKT()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CComBSTR bstrWKT("");
	CString strResult;
	VARIANT_BOOL vb;

	CComPtr<IShape> shp;
	ComHelper::CreateShape(&shp);

	int pointCount = this->GetMeasuringBase()->GetPointCount();
	// bail out if no points
	if (pointCount == 0)
	{
		this->GetMeasuring()->Clear();
		return bstrWKT;
	}

	// special handling of custom Circle measuring (type = 2)
	if (_MeasuringType == 2)
	{
		long idx;
		// create Polygon Shape
		shp->Create(ShpfileType::SHP_POLYGON, &vb);
		// radians conversion
		double radians = (2.0 * 3.14159265359 / 90.0);
		// center point
		double x, y, x2, y2;
		MeasurePoint* mp = GetMeasuringBase()->GetPoint(0);
		x = mp->x;
		y = mp->y;
		mp = GetMeasuringBase()->GetPoint(1);
		x2 = mp->x;
		y2 = mp->y;
		// we have to transform points to the map projection, build circle, then transform back
		_WGS84->Transform(&x, &y, &vb);
		_WGS84->Transform(&x2, &y2, &vb);
		// radius
		double radius = sqrt((x2 - x) * (x2 - x) + (y2 - y) * (y2 - y));
		// circle will be a clockwise polygon of points
		for (int i = 0; i <= 90; i++)
		{
			double newX = x + radius * cos(i * radians);
			double newY = y - radius * sin(i * radians);
			_MapProj->Transform(&newX, &newY, &vb);
			shp->AddPoint(newX, newY, &idx);
		}
		// reset the last point to the first point, since mathematically, they are not quite equal
		shp->get_XY(0, &x, &y, &vb);
		shp->put_XY(90, x, y, &vb);
	}
	else if (pointCount == 1 || _MeasuringType == 3)
	{
		long idx;
		// custom point measuring type
		shp->Create(ShpfileType::SHP_POINT, &vb);
		// get one and only point
		double x, y;
		MeasurePoint* mp = GetMeasuringBase()->GetPoint(0);
		x = mp->x;
		y = mp->y;
		shp->AddPoint(x, y, &idx);
	}
	else // standard measuring, ask the measure what it is
	{
		// else use measuring data
		bool isPolygon = false;
		bool isClosedLine = false;
		// what do we have?
		if (this->GetMeasuringBase()->HasPolygon(false))
		{
			shp->Create(ShpfileType::SHP_POLYGON, &vb);
			isPolygon = true;
		}
		else if (this->GetMeasuringBase()->HasLine(false) && pointCount > 1)
		{			
			// watch for last point ~equal to first point
			// for which we'll treat it like a closed polygon
			if (pointCount >= 4)
			{
				MeasurePoint* mpFirst = GetMeasuringBase()->GetPoint(0);
				MeasurePoint* mpLast = GetMeasuringBase()->GetPoint(pointCount - 1);
				// we need the points in the Map projection to get the proper distance
				double x, y, x2, y2;
				x = mpFirst->x; y = mpFirst->y;
				x2 = mpLast->x; y2 = mpLast->y;
				_WGS84->Transform(&x, &y, &vb);
				_WGS84->Transform(&x2, &y2, &vb);
				double distance = sqrt(pow(x2 - x, 2) + pow(y2 - y, 2));
				double snapTolerance = GetMouseTolerance(ToleranceSnap, true);
				// if point distance is within snap tolerance...
				if (distance <= snapTolerance)
				{
					// consider it a polygon
					mpLast->x = mpFirst->x;
					mpLast->y = mpFirst->y;
					shp->Create(ShpfileType::SHP_POLYGON, &vb);
					isClosedLine = true;
				}
			}
			// unless we decided to make it a polygon...
			if (!isClosedLine)
			{
				shp->Create(ShpfileType::SHP_POLYLINE, &vb);
			}
			// polygon is FALSE indicating that we don't add a point below to close a polygon
			isPolygon = false;
		}
		else
			// nothing, return empty WKT
			return bstrWKT;

		// iterate points in shape
		for (int idx = 0; idx < pointCount; idx++)
		{
			long newIdx;
			double x, y;
			// transfer to shape points
			MeasurePoint* pPoint = this->GetMeasuringBase()->GetPoint(idx);
			shp->AddPoint(pPoint->x, pPoint->y, &newIdx);
			// if on the first point...
			if (idx == 0)
			{
				// save the first point
				x = pPoint->x; y = pPoint->y;
			}
			// if on the last point of a polygon...
			if (idx == (pointCount - 1) && isPolygon)
			{
				// we don't know if the polygon came from the line drawing or the area drawing,
				// but if it came from the area, the last point is not equal to the first point
				if (pPoint->x != x || pPoint->y != y)
				{
					// add new last point, being the same as first point
					shp->AddPoint(x, y, &newIdx);
				}
			}
		}
	}

	// get WKT
	shp->ExportToWKT(&bstrWKT);
	
	// PROGRESS IDE cannot accept more than 32K string
#ifdef RELEASE_MODE
	if (bstrWKT.Length() > 32000)
		bstrWKT = "NOK";
#endif // RELEASE_MODE

	// return result
	//return bstrWKT;
	// large BSTR's don't always return properly,
	// but seem to work if set into CStrings and allocated out the door
	CAtlString cstr(bstrWKT);
	return cstr.AllocSysString();
}


// ***************************************************************
//		GetMeasureRadius()
// ***************************************************************
double CMapView::GetMeasureRadius()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	long errCode;
	// for custom Circle measure, the first segment is the radius
	return GetMeasuringBase()->GetSegmentLength(0, errCode);
}


// ***************************************************************
//		GetMeasurePoint()
// ***************************************************************
void CMapView::GetMeasurePoint(DOUBLE* pX, DOUBLE* pY)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// initialize for error
	*pX = 0.0;
	*pY = 0.0;

	// we only need the zeroeth point
	MeasurePoint* mp = GetMeasuringBase()->GetPoint(0);
	
	if (mp)
	{
		*pX = mp->x;
		*pY = mp->y;

		VARIANT_BOOL vb;
		// transform to map coordinates
		_WGS84->Transform(pX, pY, &vb);
	}
}


// ***************************************************************
//		SetWKTBuffer()
// ***************************************************************
BSTR CMapView::SetWKTBuffer(LPCTSTR WKT, DOUBLE BufferSize, LONG Resolution)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CComBSTR bstrWKT(WKT);
	VARIANT_BOOL vb;

	// create geometry from incoming WKT
	CComPtr<IShape> shp;
	ComHelper::CreateShape(&shp);
	shp->ImportFromWKT(bstrWKT, &vb);
	ShpfileType sft;
	shp->get_ShapeType(&sft);

	// transform to Map projection
	long count;
	double x, y;
	shp->get_NumPoints(&count);
	for (long i = 0; i < count; i++)
	{
		shp->get_XY(i, &x, &y, &vb);
		_WGS84->Transform(&x, &y, &vb);
		shp->put_XY(i, x, y, &vb);
	}

	// get buffer size in map units
	tkUnitsOfMeasure mapUnits;
	_MapProj->get_LinearUnits(&mapUnits);
	// convert incoming meters to Map units
	GetUtils()->ConvertDistance(tkUnitsOfMeasure::umMeters, mapUnits, &BufferSize, &vb);

	//CComPtr<IShape> tempShp;
	//shp->Buffer(0.01, 16, &tempShp);
	// add buffer in proper units
	CComPtr<IShape> bufferShp;
	shp->Buffer(BufferSize, Resolution, &bufferShp);

	// point count has changed
	bufferShp->get_NumPoints(&count);
	// transform back to WGS84 projection
	for (long i = 0; i < count; i++)
	{
		bufferShp->get_XY(i, &x, &y, &vb);
		_MapProj->Transform(&x, &y, &vb);
		bufferShp->put_XY(i, x, y, &vb);
	}

	// export back to WKT
	CComBSTR bstrResult;
	bufferShp->ExportToWKT(&bstrResult);

	// PROGRESS IDE cannot accept more than 32K string
#ifdef  RELEASE_MODE
	if (bstrResult.Length() > 32000)
		bstrResult = "NOK";
#endif //  RELEASE_MODE

	//return bstrResult;
	// large BSTR's don't always return properly,
	// but seem to work if set into CStrings and allocated out the door
	CAtlString cstr(bstrResult);
	return cstr.AllocSysString();
}


// ***************************************************************
//		SetWKTBufferEx()
// ***************************************************************
BSTR CMapView::SetWKTBufferEx(LPCTSTR WKT, DOUBLE BufferSize, LONG Resolution, BOOL SingleSided, LONG CapStyle, LONG JoinStyle, DOUBLE MitreLimit)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CComBSTR bstrWKT(WKT);
	VARIANT_BOOL vb;

	// create geometry from incoming WKT
	CComPtr<IShape> shp;
	ComHelper::CreateShape(&shp);
	shp->ImportFromWKT(bstrWKT, &vb);
	ShpfileType sft;
	shp->get_ShapeType(&sft);

	// transform to Map projection
	long count;
	double x, y;
	shp->get_NumPoints(&count);
	for (long i = 0; i < count; i++)
	{
		shp->get_XY(i, &x, &y, &vb);
		_WGS84->Transform(&x, &y, &vb);
		shp->put_XY(i, x, y, &vb);
	}

	// get buffer size in map units
	tkUnitsOfMeasure mapUnits;
	_MapProj->get_LinearUnits(&mapUnits);
	// convert incoming meters to Map units
	GetUtils()->ConvertDistance(tkUnitsOfMeasure::umMeters, mapUnits, &BufferSize, &vb);

	// add buffer in proper units
	CComPtr<IShape> bufferShp;
	shp->BufferWithParams(BufferSize, Resolution, SingleSided, (tkBufferCap)CapStyle, (tkBufferJoin)JoinStyle, MitreLimit, &bufferShp);

	// point count has changed
	bufferShp->get_NumPoints(&count);
	// transform back to WGS84 projection
	for (long i = 0; i < count; i++)
	{
		bufferShp->get_XY(i, &x, &y, &vb);
		_MapProj->Transform(&x, &y, &vb);
		bufferShp->put_XY(i, x, y, &vb);
	}

	// export back to WKT
	CComBSTR bstrResult;
	bufferShp->ExportToWKT(&bstrResult);

	// PROGRESS IDE cannot accept more than 32K string
#ifdef RELEASE_MODE
	if (bstrResult.Length() > 32000)
		bstrResult = "NOK";
#endif // RELEASE_MODE

	//return bstrResult;
	// large BSTR's don't always return properly,
	// but seem to work if set into CStrings and allocated out the door
	CAtlString cstr(bstrResult);
	return cstr.AllocSysString();
}


// *************************************************
//			WriteSnapshotToDC()						  
// *************************************************
void CMapView::WriteSnapshotToDC(LONG hDC, LONG WidthInPixels)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// write current map image to the specified DC
	SnapShotToDC((PVOID)hDC, GetExtents(), WidthInPixels);
}


#pragma region EnumDisplayMonitors

BOOL CALLBACK rectEnumProc(HMONITOR hMonitor, HDC hDC, LPRECT lpRECT, LPARAM dwData)
{
	//reinterpret_cast< std::vector<RECT>* >(dwData)->push_back(*lpRECT);
	reinterpret_cast< std::vector<LONG>* >(dwData)->push_back((LONG)hMonitor);
	return true;
}

BOOL CALLBACK monitorEnumProc(HMONITOR hMonitor, HDC hDC, LPRECT lpRECT, LPARAM dwData)
{
	MONITORINFOEX monitorInfo;
	monitorInfo.cbSize = sizeof(MONITORINFOEX);
	GetMonitorInfo(hMonitor, &monitorInfo);

	if (monitorInfo.dwFlags == DISPLAY_DEVICE_MIRRORING_DRIVER)
	{
		return true;
	}
	else
	{
		reinterpret_cast< std::vector<MONITORINFOEX>* >(dwData)->push_back(monitorInfo);
		return true;
	};
}

std::vector<MONITORINFOEX> monitorArray;
std::vector<RECT> rectArray;
std::vector<LONG> longArray;
// ***************************************************************
//		EnumerateDisplays()
// ***************************************************************
BSTR CMapView::EnumerateDisplays()
{

	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CString strResult;

	// initiate the call
	monitorArray.clear();
	rectArray.clear();
	longArray.clear();
	//if (EnumDisplayMonitors(NULL, NULL, monitorEnumProc, reinterpret_cast<LPARAM>(&monitorArray)))
	//if (EnumDisplayMonitors(NULL, NULL, rectEnumProc, reinterpret_cast<LPARAM>(&rectArray)))
	if (EnumDisplayMonitors(NULL, NULL, rectEnumProc, reinterpret_cast<LPARAM>(&longArray)))
		{
		strResult = "";
		//for (MONITORINFOEX monitorInfo : monitorArray)
		//for (RECT rect : rectArray)
		for (LONG hMon : longArray)
			{
			// each display separated by ch30;
			// if we're here and string already has characters, than we're on the next display
			if (strResult.GetLength() > 0) strResult.Append(ch30);

			//// left, top, right, bottom, separated by ch31
			//strResult.Format("%d%s%d%s%d%s%d", rect.left, ch31, rect.top, ch31, rect.right, ch31, rect.bottom);


			// hMonitor values, separated by ch30
			strResult.Format("%d", hMon);
		}
	}

	return strResult.AllocSysString();
}

#pragma endregion


// ***************************************************************
//		SetLayerGeomSelection()
// ***************************************************************
void CMapView::SetLayerGeomSelection(LONG LayerHandle, LPCTSTR GeometryHandles, BOOL Selected)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// get Shapefile reference
	CComPtr<IShapefile> sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// parse geometry handles
		std::vector<CAtlString> strHandles = ParseDelimitedStrings(GeometryHandles, ch30);
		// iterate
		for (CAtlString strHandle : strHandles)
		{
			// convert to numeric
			long handle = atol((LPCTSTR)strHandle);
			// set selection
			sf->put_ShapeSelected(handle, (Selected == FALSE) ? VARIANT_FALSE : VARIANT_TRUE);
		}
	}
}

// ***************************************************************
//		SetLayerSelectionColor()
// ***************************************************************
void CMapView::SetLayerSelectionColor(LONG LayerHandle, LONG RgbColor)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// get Shapefile reference
	CComPtr<IShapefile> sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// layer should be 'selectable'
		sf->put_Selectable(VARIANT_TRUE);
		// set selection appearance to 'color'
		sf->put_SelectionAppearance(tkSelectionAppearance::saSelectionColor);
		// set the color
		sf->put_SelectionColor(RgbColor);
	}
}

// ***************************************************************
//		SetLayerSelectionTransparency()
// ***************************************************************
void CMapView::SetLayerSelectionTransparency(LONG LayerHandle, DOUBLE PercentTransparency)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// get Shapefile reference
	CComPtr<IShapefile> sf = GetShapefile(LayerHandle);
	if (sf)
	{
		// layer should be 'selectable'
		sf->put_Selectable(VARIANT_TRUE);
		// set selection appearance to 'color'
		sf->put_SelectionAppearance(tkSelectionAppearance::saSelectionColor);
		// transparency is a value 0-255, where 0 is transparent
		BYTE bTransparency = static_cast<BYTE>(255.0 - (255.0 * (PercentTransparency / 100.0)));
		sf->put_SelectionTransparency(bTransparency);
	}
}

