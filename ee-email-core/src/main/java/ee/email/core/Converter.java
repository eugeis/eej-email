/*
 * Controlguide
 * Copyright (c) Siemens AG 2015, All Rights Reserved, Confidential
 */
package ee.email.core;

public interface Converter<F, T> {
  T convert(F from);
}
