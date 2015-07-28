/*
 * Copyright 2015-2015 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package ee.email.core;

import java.util.Set;

public class SetFolderFilter implements FolderFilter {

  private final Set<String> folders;

  private boolean include = true;

  public SetFolderFilter(Set<String> folders, boolean include) {

    super();
    this.folders = folders;
    this.include = include;
  }

  @Override
  public boolean isFolderToParse(String folderPath) {

    boolean ret;
    if (this.include) {
      ret = this.folders.contains(folderPath);
    } else {
      ret = !this.folders.contains(folderPath);
    }
    return ret;
  }

}
