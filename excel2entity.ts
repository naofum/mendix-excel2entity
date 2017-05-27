/*
    Copyright (C) 2017 Naofumi

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

import { MendixSdkClient, OnlineWorkingCopy, Project, Revision, Branch, loadAsPromise } from "mendixplatformsdk";
import { ModelSdkClient, IModel, projects, domainmodels, microflows, pages, navigation, texts, security, IStructure, menus } from "mendixmodelsdk";
import when = require('when');
import XLSX = require('xlsx');
import {IWorkSheet, IWorkBook} from "xlsx";
const utils = XLSX.utils;

const username = `{YOUR_USERNAME}`;
const apikey = `{YOUR_API_KEY}`;
const projectName = `{YOUR_PROJECT_NAME}`;
const projectId = `{YOUR_PROJECT_ID}`;
const moduleName = "MyFirstModule";
const revNo = -1; // -1 for latest
const branchName = null // null for mainline
const wc = null;
const client = new MendixSdkClient(username, apikey);
const project = new Project(client, projectId, projectName);
const revision = new Revision(revNo, new Branch(project, branchName));

client.platform().createOnlineWorkingCopy(project, revision)
    .then(workingCopy => loadDomainModel(workingCopy))
    .then(workingCopy => {
        const dm = pickDomainModel(workingCopy);
        const domainModel = dm.load();
        createEntities(domainModel);
        return workingCopy;
    })
    .then(workingCopy => workingCopy.commit())
    .done(
        revision => console.log(`Successfully committed revision: ${revision.num() }. Done.`),
        error => {
            console.log('Something went wrong:');
            console.dir(error);
        });

function loadDomainModel(workingCopy: OnlineWorkingCopy): when.Promise<OnlineWorkingCopy> {
    const dm = pickDomainModel(workingCopy);
    return when.promise<OnlineWorkingCopy>((resolve, reject) => {
        dm.load(dm => resolve(workingCopy));
    });
}

function pickDomainModel(workingCopy: OnlineWorkingCopy): domainmodels.IDomainModel {
    return workingCopy.model().allDomainModels()
        .filter(dm => dm.qualifiedName === moduleName)[0];
}

function createEntities(domainModel: domainmodels.DomainModel): when.Promise<void> {
    var workbook = XLSX.readFile('template.xlsx');
    var sheet_name_list = workbook.SheetNames;
    var xLoc = 100;
    var yLoc = 100;
    sheet_name_list.forEach(function (sname) {
      var worksheet = workbook.Sheets[sname];
      var cell_a1 = worksheet['A1'];
      if (((cell_a1 ? cell_a1.v : undefined) == 'Table name (logical name)') &&
          (sname != 'all attributes') &&
          (sname != 'Tables in no category') &&
          (sname != 'all tables')) {
        var entity = domainmodels.Entity.createIn(domainModel);
        entity.name = camelCase((worksheet['B2'] ? worksheet['B2'].v : '').toLowerCase());
        entity.documentation = (worksheet['B1'] ? worksheet['B1'].v : '');
        entity.location = { x: xLoc, y: yLoc };
        var range = utils.decode_range(worksheet['!ref']);
        for(var R = 7; R <= range.e.r; ++R) {
          if (!worksheet['A' + R]) {
            break;
          }
          var attr = domainmodels.Attribute.createIn(entity);
          var type = (worksheet['C' + R] ? worksheet['C' + R].v : '');
          var len = (worksheet['D' + R] ? worksheet['D' + R].v : '');
          attr.name = camelCase((worksheet['B' + R] ? worksheet['B' + R].v : '').toLowerCase());
          if (type.lastIndexOf('char', 0) === 0) {
            var stringAttributeType = domainmodels.StringAttributeType.createIn(attr);
            stringAttributeType.length = len;
            attr.type = stringAttributeType;
          } else if (type.lastIndexOf('varchar', 0) === 0) {
            var stringAttributeType = domainmodels.StringAttributeType.createIn(attr);
            stringAttributeType.length = len;
            attr.type = stringAttributeType;
          } else if (type.lastIndexOf('number', 0) === 0) {
            var decimalAttributeType = domainmodels.DecimalAttributeType.createIn(attr);
            attr.type = decimalAttributeType;
          } else if (type.lastIndexOf('timestamp', 0) === 0) {
            var dateTimeAttributeType = domainmodels.DateTimeAttributeType.createIn(attr);
            attr.type = dateTimeAttributeType;
          } else {
            var stringAttributeType = domainmodels.StringAttributeType.createIn(attr);
            attr.type = stringAttributeType;
          }
          attr.documentation = (worksheet['A' + R] ? worksheet['A' + R].v : '');
        }
        xLoc += 50;
        yLoc += 50;
      }
    });
    return;
}

function camelCase(str){
    str = str.charAt(0).toLowerCase() + str.slice(1);
    return str.replace(/[-_](.)/g, function(match, group1) {
        return group1.toUpperCase();
    });
}
