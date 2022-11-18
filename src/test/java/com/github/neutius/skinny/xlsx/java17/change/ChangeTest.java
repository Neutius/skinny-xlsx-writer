package com.github.neutius.skinny.xlsx.java17.change;

import org.junit.jupiter.api.Test;

import java.time.LocalDate;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class ChangeTest {

	@Test
	void created_nullValues() {
		Change created = new Created(null, null);

		assertThat(created.getInitialObject()).isNull();
		assertThat(created.getNewObject()).isNull();
	}

	@Test
	void validatedCreated_nullValues() {
		assertThatThrownBy(() -> new ValidatedCreated(null, null))
				.isInstanceOf(IllegalArgumentException.class)
				.hasMessageContainingAll("created", "null");
	}

	@Test
	void validatedCreated_properValues() {
		LocalDate newDate = LocalDate.of(1973, 7, 16);
		Change created = new Created(null, newDate);
		Change validatedCreated = new ValidatedCreated(null, newDate);

		assertThat(created.getInitialObject()).isNull();
		assertThat(validatedCreated.getInitialObject()).isNull();
		assertThat(created.getNewObject()).isEqualTo(validatedCreated.getNewObject());
	}

}
