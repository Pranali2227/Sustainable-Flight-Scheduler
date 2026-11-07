import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { PresentationFile } from '@oai/artifact-tool';

const root='C:/Users/patil/OneDrive/Desktop/College DA/Sem 3/DBMS/Project';
const skill='C:/Users/patil/.codex/plugins/cache/openai-primary-runtime/presentations/26.1011.11113/skills/presentations';
const build=path.join(root,'.codex-build');
const candidatePath=path.join(build,'draft','candidate.pptx');
const finalPath=path.join(root,'deliverables','Airline_Logistics_Sustainability_Presentation_Final.pptx');
await fs.mkdir(path.dirname(finalPath),{recursive:true});
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const result=await finalizePresentation({
  explicitTotalSlideCount:30,
  requiredNativeTableOwnerSlides:[],
  requiredNativeChartOwnerSlides:[],
  workspaceDir:root,
  candidatePath,
  finalPath,
  pythonExecutable:'C:/Users/patil/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',
  integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-heading-fit'],
  fontPolicy:{basis:'design',families:['Aptos']},
  verifyArtifactToolImport:true,
  receiptPath:path.join(build,'validation-final.json'),
});
console.log(JSON.stringify(result));
