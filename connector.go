// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import "fmt"

// GetConnectors provides a function to get connectors (p:cxnSp shapes) by
// given slide id.
func (f *File) GetConnectors(slideID int) ([]decodeConnector, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	return slide.getConnectors(), nil
}

// getConnectors returns connectors of the slide.
func (ds *decodeSlide) getConnectors() []decodeConnector {
	return ds.CommonSlideData.ShapeTree.Connector
}

// GetConnector provides a function to get a connector by given slide id and
// connector id (the id attribute of the p:cNvPr element).
func (f *File) GetConnector(slideID int, connectorID int) (*decodeConnector, error) {
	connectors, err := f.GetConnectors(slideID)
	if err != nil {
		return nil, err
	}

	for i := range connectors {
		if connectors[i].NonVisualConnectorProperties != nil &&
			connectors[i].NonVisualConnectorProperties.CommonNonVisualProperties != nil &&
			connectors[i].NonVisualConnectorProperties.CommonNonVisualProperties.ID == connectorID {
			return &connectors[i], nil
		}
	}

	return nil, ErrConnectorNotExist{connectorID}
}

// CreateConnector provides the function to create a new connector
// (p:cxnSp shape) by given slide id and returns the id of the connector in
// the slide after it was appended.
func (f *File) CreateConnector(slideID int, shapeProperties DecodeShapeProperties) (int, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return -1, err
	}

	connectors := slide.getConnectors()

	// the connector id must be unique in the whole shape tree (sp + pic +
	// cxnSp), the same way CreateShape and AddImage allocate shape ids
	connectorID := f.nextShapeID(slide)

	name := fmt.Sprintf("Connector %d", connectorID)
	newConnector := decodeConnector{
		NonVisualConnectorProperties: &decodeNonVisualConnectorProperties{
			CommonNonVisualProperties: &decodeCommonNonVisualProperties{
				ID:   connectorID,
				Name: name,
			},
			// the connector non-visual properties and the non-visual
			// properties elements are expected by the consumers (the same
			// way CreateShape writes the empty elements of the shape)
			CommonNonVisualConnectorProperties: &CommonNonVisualConnectorProperties{},
			NonVisualProperties:                &decodeNonVisualProperties{},
		},
		ShapeProperties: &shapeProperties,
	}

	connectors = append(connectors, newConnector)

	slide.CommonSlideData.ShapeTree.Connector = connectors

	return connectorID, nil
}

// DeleteConnector provides a function to delete a connector on the slide by
// given slide id and connector id (the id attribute of the p:cNvPr element).
func (f *File) DeleteConnector(slideID int, connectorID int) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	connectors := slide.getConnectors()

	deleteIndex := -1
	for i := range connectors {
		if connectors[i].NonVisualConnectorProperties != nil &&
			connectors[i].NonVisualConnectorProperties.CommonNonVisualProperties != nil &&
			connectors[i].NonVisualConnectorProperties.CommonNonVisualProperties.ID == connectorID {

			deleteIndex = i

			break
		}
	}

	if deleteIndex == -1 {
		return ErrConnectorNotExist{connectorID}
	}

	connectors = append(connectors[:deleteIndex], connectors[deleteIndex+1:]...)

	slide.CommonSlideData.ShapeTree.Connector = connectors

	return nil
}
