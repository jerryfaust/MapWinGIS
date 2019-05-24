#include "stdafx.h"
#include "VertexEditor.h"
#include "ShapeEditor.h"
#include "ShapeHelper.h"

// ************************************************************
//		StartEdit
// ************************************************************
void VertexEditor::StartEdit(CShapeEditor* editor, long layerHandle, long shapeIndex)
{
	editor->Clear();
	if (layerHandle != -1 && shapeIndex != -1)
	{
		VARIANT_BOOL vb;
		editor->StartEdit(layerHandle, shapeIndex, &vb);
		if (vb) {
			editor->SetRedrawNeeded(rtVolatileLayer);
		}
	}
}

// ************************************************************
//		OnMouseDown
// ************************************************************
bool VertexEditor::OnMouseDown(IMapViewCallback* map, CShapeEditor* editor, double projX, double projY, bool ctrl, bool shift)
{
	if (!editor) return true;

	VARIANT_BOOL isEmpty;
	editor->get_IsEmpty(&isEmpty);

	long layerHandle = -1, shapeIndex = -1;
	if (!isEmpty)
	{
		if (ctrl)
		{
			bool stopped = editor->TryStop();
			if (stopped)
				editor->SetRedrawNeeded(rtVolatileLayer);
			return true;
		}
		
		if (!OnMouseDownEditing(map, editor, projX, projY, shift))
		{
			if (!editor->TryStop())   // it was clicked outside shape; so clear it
				return true;
		}
		else 
			return true;
	}
	return false;	
}

// ************************************************************
//		OnMouseDownEditing
// ************************************************************
bool VertexEditor::OnMouseDownEditing(IMapViewCallback* map, CShapeEditor* editor, double projX, double projY, bool shift)
{
	CComPtr<IShape> shp = NULL;
	editor->get_RawData(&shp);

	double tol = map->_GetMouseProjTolerance();

	tkEditorBehavior behavior;
	editor->get_EditorBehavior(&behavior);

	// select vertex
	if (behavior == ebVertexEditor)
	{
		int pntIndex = editor->GetClosestVertex(projX, projY, tol);
		if (pntIndex != -1)
		{
			// start vertex moving
			bool changed = editor->SetSelectedVertex(pntIndex);
			map->_StartDragging(DragMoveVertex);
			editor->SaveState();
			if (changed) editor->SetRedrawNeeded(rtShapeEditor);
			return true;
		}
	}

	// select part
	if (behavior == ebPartEditor)
	{
		double x, y;
		if (editor->GetClosestPoint(projX, projY, x, y))
		{
			double dist = sqrt(pow(x - projX, 2.0) + pow(y - projY, 2.0));
			if (dist < tol)
			{
				// select part
				int partIndex = editor->SelectPart(x, y);
				if (partIndex != -1)
				{
					if (editor->SetSelectedPart(partIndex)) {
						map->_StartDragging(DragMovePart);
					}
					editor->SetRedrawNeeded(rtShapeEditor);
					return true;
				}
				
				if (editor->HasSelectedPart()) 
				{
					map->_StartDragging(DragMovePart);
					return true;
				}
			}
		}
	}

	// MWGIS-153: until further review, due to difficulty in editing, 
	// disable 'moving' altogether within the context of the vertex editor
	//
	//// start shape moving (only if 'shift' is pressed)
	//if (shift && ShapeHelper::PointWithinShape(shp, projX, projY, tol))
	//{
	//	// it's confusing to have both part and shape move depending on where you clicked
	//	if (editor->HasSelectedPart()) return true;

	//	map->_StartDragging(DragMoveShape);
	//	return true;
	//}
	return false;
}