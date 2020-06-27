package net.codejava.repository;

import org.springframework.data.repository.CrudRepository;
import org.springframework.stereotype.Repository;

import net.codejava.model.Employee;

@Repository
public interface EmployRepository extends CrudRepository<Employee, Long> {

}
