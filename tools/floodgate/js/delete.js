import { getConfig as getFloodgateConfig } from './config.js';
import { validateConnection, deleteFile, updateExcelTable, getAuthorizedRequestOption } from '../../loc/sharepoint.js';
import { initProject as getProjectFile, PROJECT_STATUS } from './project.js';
import { loadingON } from '../../loc/utils.js';
import { updateProjectStatusUI } from './ui.js';

async function deleteItemInProject(filePath) {
  const status = { success: false };
  let deleteSuccess = true;
  try {
    validateConnection();
    const { sp } = await getFloodgateConfig();
    const baseURI = sp.api.file.get.fgBaseURI;
    await deleteFile(sp, `${baseURI}${filePath}`);
    deleteSuccess = true;
  } catch (error) {
    console.log(`Error occurred when trying to delete files of main content tree ${error.message}`);
  }
  status.success = deleteSuccess;
  status.srcPath = filePath;
  return status;
}

async function deleteAll() {
  const startDelete = new Date();
  const projectFile = await getProjectFile();
  const deleteStatuses = [];
  const { sp } = await getFloodgateConfig();
  const fgFolders = [''];
  const baseURI = sp.api.file.get.fgBaseURI;
  const temp = '/drafts/nsivakum/trial';
  const slash = '/';
  const finalBaserURI = baseURI + temp;
  const options = getAuthorizedRequestOption({ method: 'GET' });
  const uri = `${finalBaserURI}${fgFolders.shift()}:/children`;
  const res = await fetch(uri, options);
  if (res.ok) {
    const json = await res.json();
    const files = json.value;
    for (let i = 0; i < files.length; i += 1) {
      const status = await deleteItemInProject(temp + slash + files[i].name);
      deleteStatuses.push(status);
    }
  }
  const endDelete = new Date();
  const failedDeletes = deleteStatuses.filter((status) => !status.success)
    .map((status) => status.srcPath || 'Path Info Not available');
  const excelValues = [['DELETE', startDelete, endDelete, failedDeletes.join('\n')]];
  const excelPath = await projectFile.excelPath;
  await updateExcelTable(excelPath, 'DELETE_STATUS', excelValues);
  loadingON('Project excel file updated with promote status... ');
  if (failedDeletes.length > 0) {
    loadingON('Error occurred when deleting floodgated content. Check project excel sheet for additional information<br/><br/>');
  } else {
    loadingON('Deleted floodgate tree successfully.');
  }
  const deleteStatus = { startTime: startDelete, type: 'deleteAction' };
  deleteStatus.status = (failedDeletes.length > 0)
    ? PROJECT_STATUS.COMPLETED_WITH_ERROR
    : PROJECT_STATUS.COMPLETED;
  deleteStatus.message = (failedDeletes.length > 0)
    ? 'Error occurred when deleting floodgated content. Check project excel sheet for additional information<br/><br/>'
    : 'Deleted floodgate tree successfully.';
  updateProjectStatusUI({ deleteStatus });
}

export { deleteAll };
